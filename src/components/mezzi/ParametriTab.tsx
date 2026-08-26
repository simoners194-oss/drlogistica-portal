// DR Portal — modulo Mezzi: tab Parametri (configurazione, export, import seed).

import { useState } from "react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Download, Save, Upload } from "lucide-react";
import { esportaCsvFile, csvData } from "@/lib/csv";
import { affidamentoCorrente, type MezziDb, type ParametriMezzi } from "@/lib/mezzi-types";
import { spMezziImportaDb, spMezziSalvaParametri } from "@/lib/mezzi.functions";
import { ConfirmButton, Field, fmtData, inputCls, useMezzi } from "./shared";

export function ParametriTab({
  db,
  onDb,
  isAdmin,
}: {
  db: MezziDb;
  onDb: (db: MezziDb) => void;
  isAdmin: boolean;
}) {
  const { m } = useMezzi();
  const [p, setP] = useState<ParametriMezzi>(db.parametri);
  const [json, setJson] = useState("");
  const [saving, setSaving] = useState(false);
  const [esitoImport, setEsitoImport] = useState<string | null>(null);
  const set = (patch: Partial<ParametriMezzi>) => setP((prev) => ({ ...prev, ...patch }));

  const salva = async () => {
    setSaving(true);
    try {
      const res = await spMezziSalvaParametri({ data: { parametri: p } });
      onDb(res);
      toast.success(m("salvato"));
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const righeParco = () => {
    const contrattoDi = (mezzoId: string) =>
      db.contratti.find((c) => c.mezzoId === mezzoId && c.attivo !== false);
    return [...db.mezzi]
      .sort((a, b) => a.targa.localeCompare(b.targa))
      .map((x) => {
        const c = contrattoDi(x.id);
        const aff = affidamentoCorrente(db.affidamenti, x.id);
        return {
          Targa: x.targa,
          Gruppo: x.gruppo,
          Stato: x.stato,
          Proprieta: x.proprieta,
          "Tipo Mezzo": x.tipoMezzo ?? "",
          Marca: x.marca ?? "",
          Modello: x.modello ?? "",
          Appalto: x.appalto ?? "",
          Autista: aff?.autistaNome ?? "",
          Noleggiatore: c?.noleggiatore ?? "",
          "Canone mese": c?.canoneMensileEur ?? "",
          "Km mese inclusi": c?.kmMeseInclusi ?? "",
          "Extra km": c?.extraKmEur ?? "",
          "Scadenza contratto": csvData(c?.dataFine),
          "Km attuali": x.kmAttuali ?? "",
          Telaio: x.libretto?.telaio ?? "",
          Telepass: x.telepass ?? "",
          Note: x.note ?? "",
        };
      });
  };

  const esportaXlsx = async () => {
    const XLSX = await import("xlsx");
    const targaDi = new Map(db.mezzi.map((x) => [x.id, x.targa]));
    const t = (id?: string) => (id ? (targaDi.get(id) ?? id) : "");
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(righeParco()), "Parco");
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        db.scadenze
          .slice()
          .sort((a, b) => (a.scadenza < b.scadenza ? -1 : 1))
          .map((s) => ({
            Targa: t(s.mezzoId),
            Tipo: s.tipo,
            Descrizione: s.descrizione ?? "",
            Scadenza: csvData(s.scadenza),
            Gestita: s.chiusa ? "SI" : "NO",
            Note: s.note ?? "",
          })),
      ),
      "Scadenze",
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        db.contratti.map((c) => ({
          Targa: t(c.mezzoId),
          Noleggiatore: c.noleggiatore ?? "",
          "Canone mese": c.canoneMensileEur ?? "",
          "+IVA": c.canonePiuIva ? "SI" : "",
          Inizio: csvData(c.dataInizio),
          Fine: csvData(c.dataFine),
          "Km mese": c.kmMeseInclusi ?? "",
          "Extra km": c.extraKmEur ?? "",
          "Preavviso gg": c.preavvisoDisdettaGiorni ?? "",
          Attivo: c.attivo === false ? "NO" : "SI",
          Note: c.note ?? "",
        })),
      ),
      "Contratti",
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        db.affidamenti.map((a) => ({
          Targa: t(a.mezzoId),
          Autista: a.autistaNome,
          Codice: a.autistaCodice ?? "",
          Appalto: a.appalto ?? "",
          Dal: csvData(a.dal),
          Al: csvData(a.al),
          Note: a.note ?? "",
        })),
      ),
      "Affidamenti",
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        db.multe.map((x) => ({
          Targa: t(x.mezzoId),
          "Data infrazione": x.dataInfrazione,
          Comune: x.comune ?? "",
          Tipo: x.tipo ?? "",
          Importo: x.importoEur ?? "",
          Autista: x.autistaNome ?? "",
          Responsabilita: x.responsabilita,
          Stato: x.stato,
          Note: x.note ?? "",
        })),
      ),
      "Multe",
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        db.ztl.map((p) => ({
          Targa: t(p.mezzoId),
          Comune: p.comune,
          Stato: p.stato,
          "Data richiesta": csvData(p.dataRichiesta),
          Dal: csvData(p.dal),
          Al: csvData(p.al),
          Protocollo: p.protocollo ?? "",
          Note: p.note ?? "",
        })),
      ),
      "ZTL",
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        db.officina.map((i) => ({
          Targa: t(i.mezzoId),
          Data: csvData(i.data),
          Tipo: i.tipo ?? "",
          Descrizione: i.descrizione ?? "",
          Fornitore: i.fornitore ?? "",
          Costo: i.costoEur ?? "",
          Km: i.km ?? "",
        })),
      ),
      "Officina",
    );
    XLSX.writeFile(wb, `Dettaglio Mezzi DR — ${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const esportaScadenze = () => {
    const targaDi = new Map(db.mezzi.map((x) => [x.id, x.targa]));
    esportaCsvFile(
      "scadenze-mezzi",
      ["Targa", "Tipo", "Descrizione", "Scadenza", "Gestita", "Note"],
      db.scadenze
        .slice()
        .sort((a, b) => (a.scadenza < b.scadenza ? -1 : 1))
        .map((s) => [
          s.mezzoId ? (targaDi.get(s.mezzoId) ?? s.mezzoId) : "",
          s.tipo,
          s.descrizione ?? "",
          csvData(s.scadenza),
          s.chiusa ? "SI" : "NO",
          s.note ?? "",
        ]),
    );
  };

  const backupJson = () => {
    const blob = new Blob([JSON.stringify(db, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `mezzi-db-backup-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const importa = async () => {
    if (!json.trim()) return;
    setSaving(true);
    setEsitoImport(null);
    try {
      const res = await spMezziImportaDb({ data: { json } });
      onDb(res);
      setJson("");
      const riepilogo = `${res.mezzi.length} mezzi · ${res.scadenze.length} scadenze · ${res.contratti.length} contratti · ${res.affidamenti.length} affidamenti · ${res.multe.length} multe · ${res.ztl.length} ZTL`;
      setEsitoImport(`✓ ${m("importato")} ${riepilogo}`);
      toast.success(`${m("importato")} ${riepilogo}`);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      setEsitoImport(`✗ ${msg}`);
      toast.error(msg);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-base">{m("tab.parametri")}</CardTitle>
        </CardHeader>
        <CardContent className="space-y-3 pt-0">
          <Field label={m("emailAlert")}>
            <input
              className={inputCls}
              value={p.emailAlert.join(", ")}
              onChange={(e) =>
                set({
                  emailAlert: e.target.value
                    .split(",")
                    .map((s) => s.trim())
                    .filter(Boolean),
                })
              }
            />
          </Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label={m("soglieAlert")}>
              <input
                className={inputCls}
                value={p.soglieAlertGiorni.join(", ")}
                onChange={(e) =>
                  set({
                    soglieAlertGiorni: e.target.value
                      .split(",")
                      .map((s) => Number(s.trim()))
                      .filter((n) => Number.isFinite(n) && n > 0),
                  })
                }
              />
            </Field>
            <Field label={m("sogliaOutlier")}>
              <input
                className={inputCls}
                value={p.sogliaOutlierPct}
                onChange={(e) => set({ sogliaOutlierPct: Number(e.target.value) || 20 })}
              />
            </Field>
            <Field label={m("extraKmDefault")}>
              <input
                className={inputCls}
                value={p.extraKmDefaultEur}
                onChange={(e) =>
                  set({ extraKmDefaultEur: Number(e.target.value.replace(",", ".")) || 0 })
                }
              />
            </Field>
            <Field label={m("costoOrario")}>
              <input
                className={inputCls}
                value={p.costoOrarioAutistaEur}
                onChange={(e) =>
                  set({ costoOrarioAutistaEur: Number(e.target.value.replace(",", ".")) || 0 })
                }
              />
            </Field>
            <Field label={m("costoCarburante")}>
              <input
                className={inputCls}
                value={p.costoCarburanteEurLitro}
                onChange={(e) =>
                  set({ costoCarburanteEurLitro: Number(e.target.value.replace(",", ".")) || 0 })
                }
              />
            </Field>
            <Field label={m("consumoMedio")}>
              <input
                className={inputCls}
                value={p.consumoMedioKmLitro}
                onChange={(e) =>
                  set({ consumoMedioKmLitro: Number(e.target.value.replace(",", ".")) || 0 })
                }
              />
            </Field>
          </div>
          <Field label={m("urlBaseDocumenti")}>
            <input
              className={inputCls}
              value={p.urlBaseDocumenti ?? ""}
              onChange={(e) => set({ urlBaseDocumenti: e.target.value })}
              placeholder="https://…sharepoint.com/…/Mezzi DR"
            />
          </Field>
          <Field label={m("causali")}>
            <textarea
              className={`${inputCls} min-h-24`}
              value={p.causaliNonConsegna.join("\n")}
              onChange={(e) =>
                set({
                  causaliNonConsegna: e.target.value
                    .split("\n")
                    .map((s) => s.trim())
                    .filter(Boolean),
                })
              }
            />
          </Field>
          <Button onClick={salva} disabled={saving}>
            <Save className="mr-1 h-4 w-4" /> {m("salva")}
          </Button>
        </CardContent>
      </Card>

      <div className="space-y-4">
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base">{m("esporta")}</CardTitle>
          </CardHeader>
          <CardContent className="flex flex-wrap gap-2 pt-0">
            <Button variant="outline" onClick={esportaXlsx}>
              <Download className="mr-1 h-4 w-4" /> Excel (.xlsx)
            </Button>
            <Button variant="outline" onClick={esportaScadenze}>
              <Download className="mr-1 h-4 w-4" /> {m("tab.scadenze")} — {m("esportaCsv")}
            </Button>
            <Button variant="outline" onClick={backupJson}>
              <Download className="mr-1 h-4 w-4" /> {m("backupJson")}
            </Button>
          </CardContent>
        </Card>

        {isAdmin && (
          <Card>
            <CardHeader className="pb-2">
              <CardTitle className="text-base">{m("importaJson")}</CardTitle>
              <p className="text-xs text-muted-foreground">{m("importaHint")}</p>
            </CardHeader>
            <CardContent className="space-y-2 pt-0">
              <textarea
                className={`${inputCls} min-h-32 font-mono text-xs`}
                value={json}
                onChange={(e) => setJson(e.target.value)}
                placeholder='{"mezzi": […], "scadenze": […], …}'
              />
              <div className="flex items-center gap-3">
                <ConfirmButton
                  label={m("importa")}
                  onConfirm={importa}
                  disabled={saving || !json.trim()}
                  size="default"
                  icon={<Upload className="mr-1 h-4 w-4" />}
                />
                {saving && (
                  <span className="text-xs text-muted-foreground">{m("importaInCorso")}</span>
                )}
              </div>
              {esitoImport && (
                <p
                  className={`rounded-lg p-2 text-xs font-medium ${
                    esitoImport.startsWith("✓")
                      ? "bg-status-present/15 text-status-present"
                      : "bg-destructive/15 text-destructive"
                  }`}
                >
                  {esitoImport}
                </p>
              )}
            </CardContent>
          </Card>
        )}

        <p className="text-xs text-muted-foreground">
          {m("aggiornatoIl")}:{" "}
          {db.aggiornatoIl ? `${fmtData(db.aggiornatoIl)} ${db.aggiornatoIl.slice(11, 16)}` : "—"}
          {db.aggiornatoDa ? ` · ${db.aggiornatoDa}` : ""} · v{db.versione}
        </p>
      </div>
    </div>
  );
}
