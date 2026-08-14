# DR Portal — scarico AUTOMATICO della PRIMA NOTA da Aruba Fatturazione.
# -----------------------------------------------------------------------------
# Fa esattamente i click di Simone (mappa del 13/08/2026):
#   login (username/PEC + password) → menu "Incassi e pagamenti" → "Prima
#   nota" → filtro anno → flag sul quadratino accanto a Data → "Seleziona
#   tutti (N)" → box "Azioni" → "Scarica Report Excel" → "Applica"
#   → download aaaammgg_ExportMovimenti.zip.  Ripetuto per ogni anno.
#
# PRIMA VOLTA:
#   pip install playwright xlrd
#   playwright install chromium
#   copia config.esempio.json in config.json e compila (credenziali SOLO qui,
#   il file e' escluso da git)
#   python scarica_aruba.py           (browser VISIBILE: guarda cosa fa)
#
# In caso di errore salva screenshot + HTML in errori/ — incollali in chat
# e correggiamo i selettori al giro successivo.
# Quando fila liscio: headless=true nel config e via di Utilita' di
# pianificazione di Windows.

import base64
import json
import urllib.parse
import re
import sys
import time
from datetime import datetime
from pathlib import Path

# Console senza UTF-8 (pipe, Utilita' di pianificazione): mai piu' crash
# sulle frecce nei messaggi di avanzamento.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

QUI = Path(__file__).parent
CONFIG = QUI / "config.json"
SCARICATI = QUI / "scaricati"
ERRORI = QUI / "errori"
RICOGNIZIONE = QUI / "ricognizione"

URL_PORTALE = "https://fatturazioneelettronica.aruba.it/"


def carica_config() -> dict:
    if not CONFIG.exists():
        esempio = {
            "username": "IL_TUO_UTENTE_O_PEC",
            "password": "LA_TUA_PASSWORD",
            "anni": [2025, 2026],
            "cron_fatture_url": "",
            "headless": False,
        }
        (QUI / "config.esempio.json").write_text(
            json.dumps(esempio, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        print(f"Manca {CONFIG}: creato config.esempio.json — copialo in config.json e compila.")
        sys.exit(1)
    return json.loads(CONFIG.read_text(encoding="utf-8"))


def dump_errore(page, nome: str) -> None:
    ERRORI.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    try:
        page.screenshot(path=str(ERRORI / f"{stamp}-{nome}.png"), full_page=True)
        (ERRORI / f"{stamp}-{nome}.html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass
    print(f"ERRORE al passo '{nome}': screenshot e HTML in {ERRORI}")


def clicca(page, descrizione: str, *selettori: str, timeout: int = 15000) -> None:
    """Prova i selettori in ordine finche' uno funziona (il sito puo' variare)."""
    print(f"  → {descrizione}")
    ultimo = None
    for sel in selettori:
        try:
            page.click(sel, timeout=timeout)
            time.sleep(1.2)
            return
        except Exception as e:  # prova il prossimo selettore
            ultimo = e
    dump_errore(page, descrizione[:40].replace(" ", "-"))
    raise RuntimeError(f"Nessun selettore ha funzionato per: {descrizione}") from ultimo


def chiudi_cookie(page) -> None:
    """Banner Cookiebot: si rifiutano i non necessari (best-effort)."""
    for sel in (
        "text=Rifiuta tutti",
        "#CybotCookiebotDialogBodyButtonDecline",
        "#CybotCookiebotDialogBodyLevelButtonLevelOptinDeclineAll",
        "text=Rifiuta",
        "text=Nega tutti",
        "text=Solo necessari",
    ):
        try:
            page.click(sel, timeout=2500)
            print("  (banner cookie chiuso)")
            time.sleep(1)
            return
        except Exception:
            continue


def clicca_bottone(page, descrizione: str, testo_regex: str, timeout: int = 15000) -> None:
    """Bottoni ExtJS: il testo e' COPERTO da un <button> trasparente che
    eredita l'etichetta accessibile — si clicca il bottone via ruolo."""
    print(f"  → {descrizione}")
    try:
        page.get_by_role("button", name=re.compile(testo_regex, re.I)).first.click(timeout=timeout)
        time.sleep(1.2)
        return
    except Exception:
        pass
    try:
        page.click(f"text=/{testo_regex}/i", timeout=5000)
        time.sleep(1.2)
        return
    except Exception:
        dump_errore(page, descrizione[:40].replace(" ", "-"))
        raise


def login(page, cfg: dict) -> None:
    print("Apro Aruba Fatturazione…")
    page.goto(URL_PORTALE)
    page.wait_for_load_state("networkidle", timeout=45000)
    chiudi_cookie(page)
    # Login Keycloak (loginfatturazione.aruba.it): campi standard.
    try:
        page.fill('input[name="username"], #username', cfg["username"], timeout=20000)
        page.fill('input[name="password"], #password', cfg["password"])
    except Exception:
        dump_errore(page, "login-campi")
        raise
    clicca(page, "premo Accedi", "#kc-login", 'button[type="submit"]', "text=Accedi")
    page.wait_for_load_state("networkidle", timeout=45000)
    print("Login inviato.")
    chiudi_cookie(page)


def scarica_prima_nota(page, anno: int) -> Path:
    print(f"[prima nota {anno}]")
    clicca(
        page,
        "apro Incassi e pagamenti (menu laterale)",
        "text=Incassi e pagamenti",
    )
    clicca(page, "apro Prima nota", "text=Prima nota")
    page.wait_for_load_state("networkidle", timeout=30000)
    # Filtro anno: e' un BOTTONE ExtJS "Anno: XXXX" che apre un menu.
    print(f"  → imposto il filtro anno {anno}")
    try:
        etichetta = page.get_by_text(re.compile(r"Anno:\s*\d{4}")).first
        attuale = etichetta.inner_text(timeout=10000)
        if str(anno) not in attuale:
            # il testo e' COPERTO dal bottone ExtJS: si clicca il bottone,
            # che eredita l'etichetta accessibile "Anno: XXXX".
            page.get_by_role("button", name=re.compile(r"Anno:")).first.click()
            time.sleep(1.2)
            # nel menu aperto l'anno e' un testo ESATTO ("2025"), cosi' non
            # si confonde con le date tipo 31/12/2025 nelle righe.
            page.get_by_text(str(anno), exact=True).first.click()
            time.sleep(2.5)
        else:
            print(f"    (gia' su {anno})")
    except Exception:
        dump_errore(page, f"filtro-anno-{anno}")
        raise
    time.sleep(1)
    chiudi_cookie(page)
    clicca(
        page,
        "flag sul quadratino accanto a Data (seleziona pagina)",
        "div.x-checkcolumn-title-wrap-el:visible",
        "div.x-checkcolumn .x-title-wrap-el:visible",
        # riserva: la spunta sulla PRIMA RIGA — basta a far comparire
        # la barra con "Seleziona tutti (N)".
        "div.x-checkcell:visible",
    )
    clicca_bottone(page, "clic su Seleziona tutti (N)", r"Seleziona tutti")
    try:
        clicca(page, "apro il box Azioni", '[placeholder="Azioni"]', "text=Azioni")
    except Exception:
        clicca_bottone(page, "apro il box Azioni (bottone)", r"Azioni")
    clicca(
        page,
        "flag su Scarica Report Excel",
        "text=Scarica Report Excel",
        "div:has-text('Scarica Report Excel'):visible",
    )
    with page.expect_download(timeout=120000) as attesa:
        clicca_bottone(page, "clic su Applica", r"Applica")
    download = attesa.value
    SCARICATI.mkdir(exist_ok=True)
    dest = SCARICATI / f"{datetime.now():%Y%m%d}-{anno}-{download.suggested_filename}"
    download.save_as(str(dest))
    print(f"  scaricato: {dest}")
    return dest


def estrai_nc_links(page, cfg: dict) -> None:
    """Estrae dal gestionale Aruba la mappa NC -> fattura rettificata
    (colonna Doc. coll.) chiamando direttamente advancedSearch con la
    sessione gia' aperta, e la spedisce a /cron-nc del portale (a blocchi).
    Il portale collega SOLO le NC ancora scollegate."""
    # La chiamata diretta senza gli header della webapp risponde 401: si apre
    # PRIMA la griglia "Fatture ricevute" e si INTERCETTANO gli header veri
    # (token di sessione compresi) della advancedSearch che fa il sito — poi
    # si riusano quelli per le nostre richieste a tutto-anno.
    # NIENTE chiamate nostre (token a firma per-richiesta, 401 garantito):
    # si naviga la griglia via UI e si CATTURANO le risposte advancedSearch
    # che fa il sito stesso — la griglia scarica l'anno intero in un colpo.
    catture: dict = {}

    def on_response(res):
        try:
            if "advancedsearch" not in res.url.lower() or res.status != 200:
                return
            m = re.search(r"services/([^/]+)/advancedSearch", res.url, re.I)
            srv = m.group(1) if m else ""
            anno_req = None
            try:
                anno_req = json.loads(res.request.post_data or "{}").get("AnnoFiscale")
            except Exception:
                pass
            corpo = res.json()
            if isinstance(corpo, dict) and corpo.get("Items") is not None:
                catture[(srv, anno_req)] = corpo
                print(f"    catturata {srv} anno {anno_req}: {len(corpo.get('Items') or [])} righe")
        except Exception:
            pass

    page.on("response", on_response)

    def attendi(srv, anno, sec=25):
        fine = time.time() + sec
        while time.time() < fine:
            if (srv, anno) in catture:
                return True
            time.sleep(0.5)
        return False

    def imposta_anno(anno):
        try:
            page.get_by_role("button", name=re.compile(r"Anno")).first.click(timeout=8000)
            time.sleep(1)
            page.get_by_text(str(anno), exact=True).first.click(timeout=8000)
            time.sleep(2.5)
            return True
        except Exception:
            print(f"    (filtro anno {anno} non impostabile: resto sull'anno di default)")
            return False

    anni_cfg = [int(a) for a in cfg.get("anni", [datetime.now().year])]
    for servizio, voce in (
        ("FatturaRicevutaFrontEnd", "Fatture ricevute"),
        ("FatturaFrontEnd", "Fatture inviate"),
    ):
        try:
            clicca(page, f"apro {voce}", f"text={voce}")
            page.wait_for_load_state("networkidle", timeout=45000)
            time.sleep(3)
        except Exception:
            print(f"  sezione {voce} non raggiungibile — salto")
            continue
        for anno in anni_cfg:
            if (servizio, anno) not in catture:
                imposta_anno(anno)
                attendi(servizio, anno, 40)
    # margine per le risposte ritardatarie, poi si legge TUTTO il catturato
    time.sleep(5)
    links = []
    for (srv, anno), dati in sorted(catture.items(), key=lambda x: (x[0][0], str(x[0][1]))):
        dire = "R" if srv.startswith("FatturaRicevuta") else "E"
        trovati = 0
        for it in dati.get("Items", []):
            tipo = str(it.get("Tipo", "")).upper()
            docs = it.get("DocumentiCollegati") or []
            fatture = [
                d for d in docs if str(d.get("Tipo", "")) == "Fattura" and d.get("Numero")
            ]
            if tipo.startswith("TD04") and fatture and it.get("SdiFileName"):
                links.append(
                    {
                        "file": it["SdiFileName"],
                        "numero": str(fatture[0]["Numero"]),
                        "dir": dire,
                    }
                )
                trovati += 1
        print(f"[collegamenti NC] {srv} {anno}: {trovati} trovati")
    visti = set()
    unici = []
    for l in links:
        if l["file"] in visti:
            continue
        visti.add(l["file"])
        unici.append(l)
    SCARICATI.mkdir(exist_ok=True)
    (SCARICATI / "nc-links.json").write_text(
        json.dumps(unici, indent=1, ensure_ascii=False), encoding="utf-8"
    )
    print(f"salvati {len(unici)} collegamenti in {SCARICATI / 'nc-links.json'}")
    base = str(cfg.get("cron_fatture_url", "")).strip()
    if not base or "/cron-fatture" not in base:
        print("cron_fatture_url non configurato nel config.json: spedizione al portale SALTATA.")
        print("  (incolla li' l'URL del sync programmato dalla card Diagnostica Aruba)")
        return
    base_nc = base.replace("/cron-fatture", "/cron-nc")
    for i in range(0, len(unici), 40):
        blocco = unici[i : i + 40]
        payload = base64.b64encode(json.dumps(blocco, ensure_ascii=False).encode("utf-8")).decode(
            "ascii"
        )
        url = base_nc + "&dati=" + urllib.parse.quote(payload, safe="")
        try:
            r = page.request.get(url, timeout=120000)
            # React SSR infila commenti tra i pezzi di testo (OK<!-- -->: …):
            # si tolgono prima di cercare l'esito.
            esito = r.text().replace("<!-- -->", "")
            i0 = esito.find("OK:")
            if i0 < 0:
                i0 = esito.find("ERRORE:")
            print(f"  blocco {i // 40 + 1}: HTTP {r.status} — {esito[i0 : i0 + 120] if i0 >= 0 else esito[:120]}")
        except Exception as e:
            print(f"  blocco {i // 40 + 1}: errore {e}")


def ricognizione_nc(page) -> None:
    """Apre Fatture ricevute CON LE ORECCHIE APERTE: registra l'indice di
    TUTTO il traffico di rete e salva i corpi delle risposte che sembrano
    dati (json o URL con invoice/fattur/search/list). Poi ricarica la
    pagina per forzare le chiamate della griglia."""
    indice = []
    catture = []

    def on_response(res):
        try:
            url = res.url
            ct = res.headers.get("content-type", "")
            indice.append(f"{res.status} {ct[:40]:40s} {url[:200]}")
            if res.status == 200 and "json" in ct and ("/services/" in url or "/api/" in url):
                try:
                    corpo = res.text()
                except Exception:
                    return
                if len(corpo) > 50:
                    dati_richiesta = ""
                    try:
                        dati_richiesta = res.request.post_data or ""
                    except Exception:
                        pass
                    catture.append((url, dati_richiesta, corpo))
        except Exception:
            pass

    page.on("response", on_response)
    clicca(page, "apro Fatture ricevute", "text=Fatture ricevute")
    page.wait_for_load_state("networkidle", timeout=45000)
    time.sleep(4)
    print("  → ricarico la pagina per forzare le chiamate dati")
    page.reload()
    page.wait_for_load_state("networkidle", timeout=45000)
    time.sleep(8)
    RICOGNIZIONE.mkdir(exist_ok=True)
    (RICOGNIZIONE / "indice.txt").write_text(chr(10).join(indice), encoding="utf-8")
    for i, (url, richiesta, corpo) in enumerate(catture[:60]):
        (RICOGNIZIONE / f"nc-{i:02d}.json").write_text(
            "URL: " + url + chr(10) + "RICHIESTA: " + richiesta[:5000] + chr(10) + chr(10) + corpo[:2_000_000],
            encoding="utf-8",
        )
    print(f"salvate {len(catture)} risposte dati + indice.txt in {RICOGNIZIONE}")


def spedisci_incassi(cfg: dict) -> None:
    """Aggrega le rate dei ReportMovimenti GIA' SCARICATI (l'ultimo zip per
    anno in scaricati/) e le spedisce a /cron-incassi. Niente browser.
    Il server applica SOLO AUMENTI: le riduzioni le conta e le ignora."""
    import urllib.request
    import zipfile
    from datetime import date, timedelta

    base = str(cfg.get("cron_fatture_url", "")).strip()
    if not base or "/cron-fatture" not in base:
        print("cron_fatture_url non configurato: spedizione incassi SALTATA.")
        return
    url_cron = base.replace("/cron-fatture", "/cron-incassi")
    per_anno = {}
    for f in sorted(SCARICATI.glob("*ExportMovimenti*.zip")):
        m = re.search(r"-(\d{4})-", f.name)
        per_anno[m.group(1) if m else "0000"] = f  # ordinati: l'ultimo vince
    if not per_anno:
        print("nessun ExportMovimenti in scaricati/: esegui prima lo scarico (modalita' default).")
        return
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

    def leggi_report(percorso):
        """Estrae le righe dal file dentro lo zip: Aruba consegna un VERO
        .xls binario (OLE2) — si legge con xlrd; se un domani tornasse
        xlsx, c'e' il ramo zip/xml. Righe come dict {colonna: testo}."""
        from xml.etree import ElementTree as ET
        import io as _io

        z = zipfile.ZipFile(percorso)
        nome = z.namelist()[0]
        dati = z.read(nome)
        if dati[:4].hex() == "d0cf11e0":  # OLE2 = .xls binario
            import xlrd

            wb = xlrd.open_workbook(file_contents=dati)
            sh = wb.sheet_by_index(0)
            righe = []
            for ri in range(sh.nrows):
                vals = {}
                for ci in range(sh.ncols):
                    cel = sh.cell(ri, ci)
                    v2 = cel.value
                    if cel.ctype == 3:  # data: seriale excel, lo capisce data_iso
                        v2 = str(v2)
                    elif isinstance(v2, float) and v2 == int(v2):
                        v2 = str(int(v2))
                    else:
                        v2 = str(v2)
                    vals[str(ci)] = v2
                righe.append(vals)
            return righe
        zz = zipfile.ZipFile(_io.BytesIO(dati))
        shared = []
        if "xl/sharedStrings.xml" in zz.namelist():
            root = ET.fromstring(zz.read("xl/sharedStrings.xml"))
            for si in root.iter(NS + "si"):
                shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
        fogli = sorted(x for x in zz.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml", x))
        root = ET.fromstring(zz.read(fogli[0]))
        righe = []
        for row in root.iter(NS + "row"):
            vals = {}
            for c in row:
                ref = c.get("r") or ""
                mcol = re.match(r"[A-Z]+", ref)
                if not mcol:
                    continue
                t = c.get("t")
                v2 = c.find(NS + "v")
                txt = v2.text if v2 is not None else ""
                if t == "s" and txt:
                    txt = shared[int(txt)]
                vals[mcol.group(0)] = txt
            righe.append(vals)
        return righe

    def norm(x):
        return re.sub(r"\s+", " ", str(x).strip().lower())

    def data_iso(v):
        v = str(v).strip()
        if re.match(r"^\d+(\.\d+)?$", v):  # seriale Excel
            return (date(1899, 12, 30) + timedelta(days=int(float(v)))).isoformat()
        m2 = re.match(r"^(\d{4})-(\d{2})-(\d{2})", v)
        if m2:
            return v[:10]
        m2 = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", v)
        if m2:
            return f"{m2.group(3)}-{int(m2.group(2)):02d}-{int(m2.group(1)):02d}"
        return ""

    aggregati = {}
    for anno, percorso in sorted(per_anno.items()):
        print(f"[incassi {anno}] {percorso.name}")
        righe = leggi_report(percorso)
        hdr = None
        mappa = {}
        for r in righe:
            normv = {k: norm(v2) for k, v2 in r.items()}
            if "numero fattura" in normv.values() and "flusso" in normv.values():
                hdr = r
                for col, v2 in normv.items():
                    mappa[v2] = col
                break
        if hdr is None:
            print("  intestazioni non riconosciute — salto questo file")
            continue
        cD = mappa.get("data")
        cC = mappa.get("cliente/fornitore") or mappa.get("cliente")
        cN = mappa.get("numero fattura")
        cF = mappa.get("flusso")
        cI = mappa.get("importo")
        oltre = False
        contate = 0
        for r in righe:
            if r is hdr:
                oltre = True
                continue
            if not oltre:
                continue
            numero = str(r.get(cN, "")).strip()
            flusso = norm(r.get(cF, "")).upper()
            if not numero or flusso not in ("INCASSO", "PAGAMENTO"):
                continue
            try:
                importo = abs(float(str(r.get(cI, "0")).replace(",", ".")))
            except ValueError:
                continue
            cliente = str(r.get(cC, "")).strip()
            k = (numero.lower(), norm(cliente), flusso)
            agg = aggregati.setdefault(
                k,
                {
                    "numero": numero,
                    "cliente": cliente,
                    "flusso": flusso,
                    "incassato": 0.0,
                    "ultimaData": "",
                },
            )
            agg["incassato"] = round(agg["incassato"] + importo, 2)
            d = data_iso(r.get(cD, ""))
            if d > agg["ultimaData"]:
                agg["ultimaData"] = d
            contate += 1
        print(f"  {contate} rate lette")
    lista = [
        {**a, "ultimaData": a["ultimaData"] or None} for a in aggregati.values() if a["incassato"] > 0
    ]
    for a in lista:
        if a["ultimaData"] is None:
            del a["ultimaData"]
    print(f"aggregati {len(lista)} totali per fattura — spedizione a blocchi da 60")
    for i in range(0, len(lista), 60):
        blocco = lista[i : i + 60]
        payload = base64.b64encode(json.dumps(blocco, ensure_ascii=False).encode("utf-8")).decode(
            "ascii"
        )
        url = url_cron + "&dati=" + urllib.parse.quote(payload, safe="")
        try:
            # Cloudflare rifiuta lo User-Agent di python-urllib (403):
            # ci si presenta come un browser qualunque.
            req = urllib.request.Request(
                url,
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) DRPortalCron/1.0"
                },
            )
            with urllib.request.urlopen(req, timeout=300) as resp:
                corpo = resp.read().decode("utf-8", "replace").replace("<!-- -->", "")
            m2 = re.search(r"(OK|ERRORE):[^<]*", corpo)
            print(f"  blocco {i // 60 + 1}: {m2.group(0)[:200] if m2 else corpo[:120]}")
        except Exception as e:
            print(f"  blocco {i // 60 + 1}: errore {e}")


def main() -> None:
    cfg = carica_config()
    # La modalita' "incassi" lavora sui file gia' scaricati: niente browser.
    if "incassi" in sys.argv:
        spedisci_incassi(cfg)
        return
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        # Smart App Control puo' bloccare il Chromium scaricato da Playwright
        # (DLL senza firma "di reputazione", errore 0x11C7): si prova prima
        # con Edge/Chrome DI SISTEMA, firmati e sempre ammessi; il Chromium
        # interno resta come ultimo ripiego.
        headless = bool(cfg.get("headless", False))
        browser = None
        ultimo_errore = None
        for canale in ("msedge", "chrome", None):
            try:
                if canale:
                    browser = p.chromium.launch(channel=canale, headless=headless)
                else:
                    browser = p.chromium.launch(headless=headless)
                print(f"Browser avviato ({canale or 'chromium interno'}).")
                break
            except Exception as e:  # noqa: BLE001 — si tenta il canale successivo
                ultimo_errore = e
        if browser is None:
            raise RuntimeError(f"Nessun browser avviabile: {ultimo_errore}")
        page = browser.new_page(accept_downloads=True)
        try:
            login(page, cfg)
            if "nclinks" in sys.argv:
                estrai_nc_links(page, cfg)
                return
            if "nc" in sys.argv:
                ricognizione_nc(page)
                return
            for anno in cfg.get("anni", [datetime.now().year]):
                scarica_prima_nota(page, int(anno))
                # si torna alla dashboard per ripartire puliti
                page.goto(URL_PORTALE)
                page.wait_for_load_state("networkidle", timeout=30000)
        finally:
            browser.close()
    print("Fatto: file in " + str(SCARICATI))


if __name__ == "__main__":
    main()
