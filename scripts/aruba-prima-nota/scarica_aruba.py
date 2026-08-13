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

import json
import sys
import time
from datetime import datetime
from pathlib import Path

QUI = Path(__file__).parent
CONFIG = QUI / "config.json"
SCARICATI = QUI / "scaricati"
ERRORI = QUI / "errori"

URL_PORTALE = "https://fatturazioneelettronica.aruba.it/"


def carica_config() -> dict:
    if not CONFIG.exists():
        esempio = {
            "username": "IL_TUO_UTENTE_O_PEC",
            "password": "LA_TUA_PASSWORD",
            "anni": [2025, 2026],
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


def login(page, cfg: dict) -> None:
    print("Apro Aruba Fatturazione…")
    page.goto(URL_PORTALE)
    page.wait_for_load_state("networkidle", timeout=45000)
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


def scarica_prima_nota(page, anno: int) -> Path:
    print(f"[prima nota {anno}]")
    clicca(
        page,
        "apro Incassi e pagamenti (menu laterale)",
        "text=Incassi e pagamenti",
    )
    clicca(page, "apro Prima nota", "text=Prima nota")
    page.wait_for_load_state("networkidle", timeout=30000)
    # Filtro anno, in alto: prima provo una select vera, poi il click su testo.
    print(f"  → imposto il filtro anno {anno}")
    impostato = False
    for sel in ("select", 'select[name*="anno" i]'):
        try:
            page.select_option(sel, label=str(anno), timeout=4000)
            impostato = True
            break
        except Exception:
            continue
    if not impostato:
        clicca(page, f"scelgo l'anno {anno}", f"text={anno}", timeout=8000)
    time.sleep(2)
    clicca(
        page,
        "flag sul quadratino accanto a Data (seleziona pagina)",
        "thead input[type=checkbox]",
        'input[type="checkbox"]',
    )
    clicca(page, "clic su Seleziona tutti", "text=Seleziona tutti")
    clicca(page, "apro il box Azioni", "text=Azioni", '[placeholder="Azioni"]')
    clicca(page, "flag su Scarica Report Excel", "text=Scarica Report Excel")
    with page.expect_download(timeout=120000) as attesa:
        clicca(page, "clic su Applica", "text=Applica")
    download = attesa.value
    SCARICATI.mkdir(exist_ok=True)
    dest = SCARICATI / f"{datetime.now():%Y%m%d}-{anno}-{download.suggested_filename}"
    download.save_as(str(dest))
    print(f"  scaricato: {dest}")
    return dest


def main() -> None:
    cfg = carica_config()
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=bool(cfg.get("headless", False)))
        page = browser.new_page(accept_downloads=True)
        try:
            login(page, cfg)
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
