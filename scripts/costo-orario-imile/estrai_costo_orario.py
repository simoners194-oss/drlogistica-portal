#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# DR Portal - hub iMile: costo del personale per ora riconosciuta dal cliente.
# ---------------------------------------------------------------------------
# Legge i prospetti presenze iMile (un file .xlsx per mese, foglio "MESE 2026")
# e il file degli stipendi mensili, e produce un workbook di analisi con le
# formule vive: ore riconosciute dal cliente, importo riconosciuto, stipendi
# pagati, costo per ora riconosciuta e margine sul personale.
#
# Uso:
#   pip install openpyxl
#   python estrai_costo_orario.py \
#       --presenze "Ore IMILE" \
#       --stipendi "STIPENDI GENNAIO-AGOSTO.xlsx" \
#       --out "costo-orario-imile.xlsx"
#
# --presenze punta a una CARTELLA: ogni .xlsx al suo interno che abbia il nome
# del mese in italiano nel filename viene letto (PRESENZE IMILE MARZO.xlsx...).
# Aggiungere un mese = copiare il nuovo file nella cartella e rilanciare.
#
# Il foglio "MESE 2026" del cliente ha questo tracciato, dalla riga 3 in poi,
# una riga per giorno di calendario:
#     A giorno  B giorno settimana  C data  D n. operatori  E ore totali
#     F ore ordinarie  G ore straordinarie  H costo ord.  I costo str.  J costo tot.
# L'ultima riga e' il "Total" del cliente: viene ignorata e i totali sono
# ricalcolati sommando i giorni, cosi' il dato e' verificato e non copiato.
#
# ATTENZIONE: i file sorgente e il workbook prodotto contengono le buste paga.
# Non vanno nel repository - vedi LEGGIMI.md accanto a questo script.
# ---------------------------------------------------------------------------

import argparse
import glob
import os
import re
import sys
from datetime import date

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MESI = ["GENNAIO", "FEBBRAIO", "MARZO", "APRILE", "MAGGIO", "GIUGNO",
        "LUGLIO", "AGOSTO", "SETTEMBRE", "OTTOBRE", "NOVEMBRE", "DICEMBRE"]

FOGLIO_MESE = "MESE 2026"

# colonne del foglio "MESE 2026" (0-based)
C_GIORNO, C_DOW, C_DATA, C_OPER, C_ORE, C_ORD, C_STR, C_COSTO_ORD, C_COSTO_STR, C_COSTO = range(10)

# ---- stile -------------------------------------------------------------
ARIAL = "Arial"
BLU = Font(name=ARIAL, size=10, color="0000FF")           # dato di input, preso dai file sorgente
NERO = Font(name=ARIAL, size=10)                           # formula
NERO_B = Font(name=ARIAL, size=10, bold=True)
TIT = Font(name=ARIAL, size=14, bold=True)
SOTTOTIT = Font(name=ARIAL, size=10, italic=True, color="595959")
INTEST = Font(name=ARIAL, size=9, bold=True, color="FFFFFF")
FILL_INTEST = PatternFill("solid", fgColor="1F3864")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")        # cella da compilare
FILL_TOT = PatternFill("solid", fgColor="D9E2F3")
FILL_FUORI = PatternFill("solid", fgColor="F2F2F2")
BORDO_SOPRA = Border(top=Side(style="thin", color="1F3864"))

F_ORE = '#,##0.00'
F_EUR = '#,##0.00\\ "€"'
F_EUR0 = '#,##0\\ "€"'
F_PCT = '0.0%'
F_INT = '#,##0'


def mese_da_nome(testo):
    t = testo.upper()
    for m in MESI:
        if m in t:
            return m
    return None


def leggi_presenze(percorso):
    """Totali mensili + righe giornaliere di un prospetto presenze."""
    wb = load_workbook(percorso, data_only=True, read_only=True)
    if FOGLIO_MESE not in wb.sheetnames:
        raise SystemExit(f"{percorso}: manca il foglio '{FOGLIO_MESE}'")
    ws = wb[FOGLIO_MESE]

    def n(v):
        return float(v) if isinstance(v, (int, float)) else 0.0

    giorni = []
    tot = dict(gg=0, ore=0.0, ord=0.0, str=0.0, dom=0.0, costo=0.0, pres=0.0, punta=0)
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not isinstance(r[C_GIORNO], (int, float)):   # salta la riga "Total" del cliente
            continue
        ore = n(r[C_ORE])
        if ore <= 0:                                     # giornata chiusa
            continue
        dow = r[C_DOW]
        dt = r[C_DATA]
        giorni.append(dict(
            data=dt.date() if hasattr(dt, "date") else dt,
            dow=dow, oper=n(r[C_OPER]), ore=ore,
            ord=n(r[C_ORD]), str=n(r[C_STR]),
            costo_ord=n(r[C_COSTO_ORD]), costo_str=n(r[C_COSTO_STR]), costo=n(r[C_COSTO]),
        ))
        tot["gg"] += 1
        tot["ore"] += ore
        tot["ord"] += n(r[C_ORD])
        tot["str"] += n(r[C_STR])
        tot["costo"] += n(r[C_COSTO])
        tot["pres"] += n(r[C_OPER])
        tot["punta"] = max(tot["punta"], int(n(r[C_OPER])))
        if dow == "Sun":
            tot["dom"] += ore
    tot["ord_feriali"] = tot["ord"] - tot["dom"]
    wb.close()
    return tot, giorni


def leggi_stipendi(percorso):
    wb = load_workbook(percorso, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    fuori = {}
    for r in ws.iter_rows(values_only=True):
        if not r or r[0] is None:
            continue
        m = mese_da_nome(str(r[0]))
        if m and len(r) > 1 and isinstance(r[1], (int, float)):
            fuori[m] = float(r[1])
    wb.close()
    if not fuori:
        raise SystemExit(f"{percorso}: nessun importo mensile riconosciuto "
                         "(atteso: colonna A = nome mese, colonna B = importo)")
    return fuori


def intesta(ws, riga, etichette, larghezze):
    for i, (t, w) in enumerate(zip(etichette, larghezze), start=1):
        c = ws.cell(row=riga, column=i, value=t)
        c.font = INTEST
        c.fill = FILL_INTEST
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[riga].height = 30


def foglio_costo(wb, mesi_ord, ore, stip, fonte_presenze, fonte_stipendi):
    ws = wb.active
    ws.title = "Costo orario"
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Hub iMile — costo del personale per ora riconosciuta dal cliente"
    ws["A1"].font = TIT
    ws["A2"] = ("Ore e importi riconosciuti dal prospetto presenze iMile, confrontati con "
                "gli stipendi mensili effettivamente pagati.")
    ws["A2"].font = SOTTOTIT

    ws["A4"] = "Carico aggiuntivo su stipendi"
    ws["A4"].font = NERO_B
    ws["B4"] = 0.0
    ws["B4"].font = BLU
    ws["B4"].fill = FILL_INPUT
    ws["B4"].number_format = F_PCT
    ws["B4"].border = Border(*[Side(style="thin", color="BF8F00")] * 4)
    ws["C4"] = ("<-- cella da compilare. 0% se la colonna Stipendi e' gia' il costo azienda pieno. "
                "Se e' la sola retribuzione, inserire qui contributi + TFR + ratei (tipicamente 25-40%).")
    ws["C4"].font = SOTTOTIT

    testate = ["Mese", "Ore riconosciute", "Riconosciuto\n€", "€/h\nriconosciuto",
               "Stipendi\n€", "Costo azienda\n€", "Costo\n€/h",
               "Margine\n€", "Margine\n€/h", "Margine\n%"]
    larghezze = [13, 17, 15, 13, 15, 15, 11, 15, 11, 10]
    R0 = 6
    intesta(ws, R0, testate, larghezze)

    prima_utile = ultima_utile = None
    r = R0 + 1
    for m in mesi_ord:
        o = ore.get(m)
        ws.cell(row=r, column=1, value=m.capitalize()).font = NERO_B
        if o:
            if prima_utile is None:
                prima_utile = r
            ultima_utile = r
            ws.cell(row=r, column=2, value=round(o["ore"], 2)).font = BLU
            ws.cell(row=r, column=3, value=round(o["costo"], 2)).font = BLU
        else:
            c = ws.cell(row=r, column=2, value="ore n/d")
            c.font = Font(name=ARIAL, size=10, italic=True, color="808080")
            c.alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=3, value=None)
            for col in range(1, 11):
                ws.cell(row=r, column=col).fill = FILL_FUORI
        ws.cell(row=r, column=5, value=round(stip.get(m, 0.0), 2)).font = BLU
        ws.cell(row=r, column=6, value=f"=E{r}*(1+$B$4)")
        ws.cell(row=r, column=4, value=f'=IF(N(B{r})=0,"",C{r}/B{r})')
        ws.cell(row=r, column=7, value=f'=IF(N(B{r})=0,"",F{r}/B{r})')
        ws.cell(row=r, column=8, value=f'=IF(N(B{r})=0,"",C{r}-F{r})')
        ws.cell(row=r, column=9, value=f'=IF(N(B{r})=0,"",D{r}-G{r})')
        ws.cell(row=r, column=10, value=f'=IF(N(C{r})=0,"",H{r}/C{r})')
        r += 1

    rt = r
    etichetta = (f"Totale {mesi_ord[0].capitalize()[:3]}–{mesi_ord[-1].capitalize()[:3]}"
                 if prima_utile is None else
                 f"Totale {ws.cell(row=prima_utile, column=1).value[:3]}"
                 f"–{ws.cell(row=ultima_utile, column=1).value[:3]}")
    ws.cell(row=rt, column=1, value=etichetta).font = NERO_B
    for col in (2, 3, 5):
        L = get_column_letter(col)
        ws.cell(row=rt, column=col, value=f"=SUM({L}{prima_utile}:{L}{ultima_utile})")
    ws.cell(row=rt, column=6, value=f"=E{rt}*(1+$B$4)")
    ws.cell(row=rt, column=4, value=f"=C{rt}/B{rt}")
    ws.cell(row=rt, column=7, value=f"=F{rt}/B{rt}")
    ws.cell(row=rt, column=8, value=f"=C{rt}-F{rt}")
    ws.cell(row=rt, column=9, value=f"=D{rt}-G{rt}")
    ws.cell(row=rt, column=10, value=f"=H{rt}/C{rt}")
    for col in range(1, 11):
        c = ws.cell(row=rt, column=col)
        c.font = NERO_B
        c.fill = FILL_TOT
        c.border = BORDO_SOPRA

    for rr in range(R0 + 1, rt + 1):
        for col in range(2, 11):
            c = ws.cell(row=rr, column=col)
            if c.font is None or c.font.color is None or c.font.color.rgb != "000000FF":
                if rr != rt:
                    c.font = NERO
            c.number_format = {2: F_ORE, 3: F_EUR0, 4: F_ORE, 5: F_EUR0,
                               6: F_EUR0, 7: F_ORE, 8: F_EUR0, 9: F_ORE, 10: F_PCT}[col]
    for rr in (rt,):
        for col in range(2, 11):
            ws.cell(row=rr, column=col).font = NERO_B

    n = rt + 2
    note = [
        "Legenda",
        "  Testo blu = dato letto dai file sorgente (non modificare a mano: si rigenera con lo script).",
        "  Testo nero = formula.  Cella gialla B4 = l'unico parametro da decidere.",
        "",
        "Come si leggono le colonne",
        "  Ore riconosciute / Riconosciuto € = somma dei giorni del prospetto presenze iMile,",
        "     ricalcolata riga per riga (foglio Giornaliero) e non copiata dal totale del cliente.",
        "  Costo azienda € = Stipendi × (1 + carico di B4).",
        "  Costo €/h = Costo azienda ÷ Ore riconosciute: e' la risposta alla domanda del proprietario.",
        "  Margine = differenza fra riconosciuto e costo del personale. E' LORDO: non toglie",
        "     attrezzature, materiali, struttura e oneri di sede.",
        "",
        "Perimetro",
        "  I mesi in grigio hanno gli stipendi ma non il prospetto presenze: senza monte ore",
        "  non producono un costo orario e restano fuori dai totali.",
        "",
        f"Fonti: presenze = {fonte_presenze}   |   stipendi = {fonte_stipendi}",
        f"Generato il {date.today().strftime('%d/%m/%Y')} da estrai_costo_orario.py",
    ]
    for i, t in enumerate(note):
        c = ws.cell(row=n + i, column=1, value=t)
        c.font = NERO_B if t and not t.startswith(" ") and not t.startswith("Fonti") and not t.startswith("Generato") else SOTTOTIT
    ws.freeze_panes = f"A{R0 + 1}"
    return ws


def foglio_dettaglio(wb, mesi_ord, ore):
    ws = wb.create_sheet("Dettaglio mensile")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Composizione del monte ore riconosciuto"
    ws["A1"].font = TIT
    ws["A2"] = "Solo i mesi con prospetto presenze disponibile."
    ws["A2"].font = SOTTOTIT

    testate = ["Mese", "Giorni\naperti", "Ore\nordinarie feriali", "Ore\ndomenicali",
               "Ore\nstraordinarie", "Ore\ntotali", "Presenze\n(giorni-uomo)",
               "Ore per\npresenza", "Punta\noperatori/giorno", "Presenze\nper giorno"]
    intesta(ws, 4, testate, [13, 10, 16, 12, 13, 12, 14, 11, 14, 12])

    r = 5
    primo = r
    for m in mesi_ord:
        o = ore.get(m)
        if not o:
            continue
        ws.cell(row=r, column=1, value=m.capitalize()).font = NERO_B
        for col, val in ((2, o["gg"]), (3, round(o["ord_feriali"], 2)), (4, round(o["dom"], 2)),
                         (5, round(o["str"], 2)), (7, int(o["pres"])), (9, o["punta"])):
            ws.cell(row=r, column=col, value=val).font = BLU
        ws.cell(row=r, column=6, value=f"=C{r}+D{r}+E{r}")
        ws.cell(row=r, column=8, value=f'=IF(G{r}=0,"",F{r}/G{r})')
        ws.cell(row=r, column=10, value=f'=IF(B{r}=0,"",G{r}/B{r})')
        r += 1
    ultimo = r - 1

    ws.cell(row=r, column=1, value="Totale").font = NERO_B
    for col in (2, 3, 4, 5, 6, 7):
        L = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"=SUM({L}{primo}:{L}{ultimo})")
    ws.cell(row=r, column=8, value=f"=F{r}/G{r}")
    ws.cell(row=r, column=9, value=f"=MAX(I{primo}:I{ultimo})")
    ws.cell(row=r, column=10, value=f"=G{r}/B{r}")
    for col in range(1, 11):
        c = ws.cell(row=r, column=col)
        c.font = NERO_B
        c.fill = FILL_TOT
        c.border = BORDO_SOPRA

    for rr in range(5, r + 1):
        for col in range(2, 11):
            c = ws.cell(row=rr, column=col)
            if rr != r and (c.font is None or c.font.color is None or c.font.color.rgb != "000000FF"):
                c.font = NERO
            c.number_format = F_INT if col in (2, 7, 9) else F_ORE

    ws.cell(row=r + 2, column=1,
            value="Le ore domenicali sono riconosciute a tariffa piena (27,50 €/h) e "
                  "alzano la tariffa media del mese.").font = SOTTOTIT
    ws.freeze_panes = "A5"


def foglio_giornaliero(wb, mesi_ord, giorni_per_mese):
    ws = wb.create_sheet("Giornaliero")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Righe giornaliere del prospetto presenze iMile"
    ws["A1"].font = TIT
    ws["A2"] = ("Base di verifica: i totali degli altri fogli sono la somma di queste righe. "
                "Le giornate chiuse (zero ore) non sono riportate.")
    ws["A2"].font = SOTTOTIT

    testate = ["Mese", "Data", "Giorno", "N.\noperatori", "Ore\ntotali", "Ore\nordinarie",
               "Ore\nstraordinarie", "Costo\nordinario €", "Costo\nstraordinario €",
               "Costo\ntotale €", "€/h\ndel giorno"]
    intesta(ws, 4, testate, [12, 12, 9, 11, 11, 11, 13, 15, 15, 14, 11])

    r = 5
    primo = r
    for m in mesi_ord:
        for g in giorni_per_mese.get(m, []):
            ws.cell(row=r, column=1, value=m.capitalize()).font = NERO
            c = ws.cell(row=r, column=2, value=g["data"])
            c.number_format = "DD/MM/YYYY"
            c.font = BLU
            ws.cell(row=r, column=3, value=g["dow"]).font = NERO
            for col, val in ((4, g["oper"]), (5, g["ore"]), (6, g["ord"]), (7, g["str"]),
                             (8, g["costo_ord"]), (9, g["costo_str"]), (10, g["costo"])):
                cc = ws.cell(row=r, column=col, value=round(val, 4))
                cc.font = BLU
                cc.number_format = F_INT if col == 4 else (F_ORE if col in (5, 6, 7) else F_EUR)
            ws.cell(row=r, column=11, value=f'=IF(E{r}=0,"",J{r}/E{r})').number_format = F_ORE
            r += 1
    ultimo = r - 1

    ws.cell(row=r, column=1, value="Totale").font = NERO_B
    for col in (4, 5, 6, 7, 8, 9, 10):
        L = get_column_letter(col)
        c = ws.cell(row=r, column=col, value=f"=SUM({L}{primo}:{L}{ultimo})")
        c.font = NERO_B
        c.number_format = F_INT if col == 4 else (F_ORE if col in (5, 6, 7) else F_EUR)
    ws.cell(row=r, column=11, value=f"=J{r}/E{r}").number_format = F_ORE
    for col in range(1, 12):
        c = ws.cell(row=r, column=col)
        c.fill = FILL_TOT
        c.border = BORDO_SOPRA
        c.font = NERO_B

    ws.auto_filter.ref = f"A4:K{ultimo}"
    ws.freeze_panes = "A5"


def foglio_note(wb, fonte_presenze, fonte_stipendi, mesi_ok, mesi_ko):
    ws = wb.create_sheet("Note e fonti")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 118

    blocchi = [
        ("T", "Metodo, fonti e cose da verificare"),
        ("S", "Come sono stati ottenuti i numeri"),
        ("P", "1. Per ogni mese si apre il prospetto presenze iMile e si legge il foglio \"MESE 2026\", "
              "che ha una riga per giorno di calendario."),
        ("P", "2. Si sommano riga per riga ore totali, ore ordinarie, straordinari, numero operatori e importo "
              "riconosciuto. La riga \"Total\" gia' presente nel file del cliente non viene usata: serve solo "
              "come controllo (su tutti i mesi verificati torna al centesimo)."),
        ("P", "3. Gli stipendi mensili arrivano dal file stipendi, un importo per mese."),
        ("P", "4. Costo per ora riconosciuta = stipendi del mese (eventualmente maggiorati del carico "
              "indicato in B4 del foglio \"Costo orario\") diviso le ore riconosciute dello stesso mese."),
        ("", ""),
        ("S", "Tariffe riconosciute da iMile (dal prospetto presenze)"),
        ("P", "   Regular hours   17,50 €/h   dalle 07:00 alle 21:59"),
        ("P", "   Overtime        20,50 €/h   dalle 22:00 alle 06:00"),
        ("P", "   Domenica        27,50 €/h   tutta la giornata"),
        ("P", "L'importo riconosciuto di ogni mese e' la somma delle ore di ciascuna fascia per la sua tariffa."),
        ("", ""),
        ("S", "Cosa va chiarito prima di presentare i numeri"),
        ("P", "A. Il file stipendi riporta un solo importo per mese, senza indicazione della sede. "
              "Tutta l'analisi assume che quell'importo sia esattamente e soltanto il personale dell'hub iMile. "
              "Se comprende altre sedi il costo orario risulta gonfiato; se esclude capi turno o "
              "amministrativi di sede, e' sottostimato."),
        ("P", "B. \"Stipendi\" non e' automaticamente \"costo azienda\". Se l'importo e' la sola retribuzione, "
              "mancano contributi, TFR e ratei di tredicesima e quattordicesima: e' esattamente a questo "
              "che serve la cella gialla B4 del foglio \"Costo orario\"."),
        ("P", "C. Il margine calcolato copre il solo costo del personale. Attrezzature, materiali di consumo, "
              "struttura e oneri di sede non sono in questi numeri."),
        ("", ""),
        ("S", "Perimetro"),
        ("P", "Mesi con ore e stipendi (entrano nei totali): " + ", ".join(m.capitalize() for m in mesi_ok)),
        ("P", "Mesi con i soli stipendi (esclusi dai totali): " +
              (", ".join(m.capitalize() for m in mesi_ko) if mesi_ko else "nessuno")),
        ("", ""),
        ("S", "File sorgente"),
        ("P", "   Presenze: " + fonte_presenze),
        ("P", "   Stipendi: " + fonte_stipendi),
        ("P", f"   Generato il {date.today().strftime('%d/%m/%Y')} con estrai_costo_orario.py "
              "(rilanciare lo script dopo aver aggiunto un mese)."),
    ]
    r = 1
    for tipo, testo in blocchi:
        c = ws.cell(row=r, column=1, value=testo)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if tipo == "T":
            c.font = TIT
        elif tipo == "S":
            c.font = NERO_B
        else:
            c.font = Font(name=ARIAL, size=10)
        if tipo == "P" and len(testo) > 105:
            ws.row_dimensions[r].height = 15 * (len(testo) // 105 + 1)
        r += 1


def main():
    ap = argparse.ArgumentParser(description="Costo orario hub iMile: ore riconosciute vs stipendi pagati.")
    ap.add_argument("--presenze", required=True, help="cartella con i file PRESENZE IMILE <MESE>.xlsx")
    ap.add_argument("--stipendi", required=True, help="file xlsx con gli stipendi mensili (mese in col. A, importo in col. B)")
    ap.add_argument("--out", default="costo-orario-imile.xlsx", help="workbook di analisi da produrre")
    a = ap.parse_args()

    if not os.path.isdir(a.presenze):
        raise SystemExit(f"--presenze deve essere una cartella: {a.presenze}")

    ore, giorni = {}, {}
    letti = []
    for p in sorted(glob.glob(os.path.join(a.presenze, "*.xlsx"))):
        if os.path.basename(p).startswith("~$"):
            continue
        m = mese_da_nome(os.path.basename(p))
        if not m:
            print(f"  ignorato (nessun mese nel nome): {os.path.basename(p)}", file=sys.stderr)
            continue
        ore[m], giorni[m] = leggi_presenze(p)
        letti.append(os.path.basename(p))
        print(f"  {m:10s} {ore[m]['gg']:2d} giorni  {ore[m]['ore']:>10,.2f} ore  "
              f"{ore[m]['costo']:>12,.2f} € riconosciuti")

    stip = leggi_stipendi(a.stipendi)
    if not ore:
        raise SystemExit("nessun prospetto presenze leggibile nella cartella indicata")

    mesi_ord = [m for m in MESI if m in ore or m in stip]
    mesi_ok = [m for m in mesi_ord if m in ore]
    mesi_ko = [m for m in mesi_ord if m not in ore]

    fonte_presenze = os.path.abspath(a.presenze) + f" ({len(letti)} file: {', '.join(letti)})"
    fonte_stipendi = os.path.abspath(a.stipendi)

    wb = Workbook()
    # le formule sono scritte senza valore in cache: si obbliga Excel/LibreOffice
    # a ricalcolare tutto all'apertura, altrimenti alcuni visualizzatori mostrano celle vuote
    wb.calculation.fullCalcOnLoad = True
    foglio_costo(wb, mesi_ord, ore, stip, fonte_presenze, fonte_stipendi)
    foglio_dettaglio(wb, mesi_ord, ore)
    foglio_giornaliero(wb, mesi_ord, giorni)
    foglio_note(wb, fonte_presenze, fonte_stipendi, mesi_ok, mesi_ko)
    wb.save(a.out)

    tot_ore = sum(ore[m]["ore"] for m in mesi_ok)
    tot_ric = sum(ore[m]["costo"] for m in mesi_ok)
    tot_stip = sum(stip.get(m, 0.0) for m in mesi_ok)
    print(f"\n  {len(mesi_ok)} mesi nel perimetro: {tot_ore:,.2f} ore  "
          f"{tot_ric:,.2f} € riconosciuti  {tot_stip:,.2f} € stipendi")
    print(f"  costo {tot_stip / tot_ore:.2f} €/h  contro {tot_ric / tot_ore:.2f} €/h riconosciuti")
    print(f"\n  scritto: {os.path.abspath(a.out)}")


if __name__ == "__main__":
    main()
