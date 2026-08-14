# DR Portal — scarico AUTOMATICO della PRIMA NOTA da Aruba Fatturazione.
# -----------------------------------------------------------------------------
# Fa esattamente i click di Simone (mappa del 13/08/2026):
#   login (username/PEC + password) → menu "Incassi e pagamenti" → "Prima
#   nota" → filtro anno → flag sul quadratino accanto a Data → "Seleziona
#   tutti (N)" → box "Azioni" → "Scarica Report Excel" → "Applica"
#   → download aaaammgg_ExportMovimenti.zip.  Ripetuto per ogni anno.
#
# PRIMA VOLTA:
#   pip install playwright
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
    links = []
    for anno in cfg.get("anni", [datetime.now().year]):
        for servizio, dire in (("FatturaRicevutaFrontEnd", "R"), ("FatturaFrontEnd", "E")):
            print(f"[collegamenti NC] {servizio} {anno}…")
            try:
                r = page.request.post(
                    URL_PORTALE + f"services/{servizio}/advancedSearch",
                    data=json.dumps({"PageNumber": 1, "PageSize": None, "AnnoFiscale": int(anno)}),
                    headers={"Content-Type": "application/json"},
                    timeout=120000,
                )
                if r.status != 200:
                    print(f"  HTTP {r.status} — salto questo servizio")
                    continue
                dati = r.json()
            except Exception as e:
                print(f"  errore: {e} — salto")
                continue
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
            print(f"  {trovati} collegamenti NC trovati")
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
            esito = r.text()
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


def main() -> None:
    cfg = carica_config()
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
