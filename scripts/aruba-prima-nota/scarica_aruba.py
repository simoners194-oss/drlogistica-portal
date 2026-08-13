# DR Portal — scarico AUTOMATICO da Aruba Fatturazione (sito web):
#   1) prima nota / ReportMovimenti (incassi e pagamenti)
#   2) elenco note di credito (per l'aggancio NC)
# -----------------------------------------------------------------------------
# L'API Aruba NON espone questi dati (verificato sulla documentazione v1/v2):
# questo script fa quello che farebbe una persona — apre il sito, entra,
# scarica gli export — e salva i file in una cartella. Le credenziali stanno
# SOLO in config.json accanto allo script (mai nel portale, mai in chat).
#
# PRIMA VOLTA:
#   pip install playwright
#   playwright install chromium
#   copia config.esempio.json in config.json e compila le credenziali
#   python scarica_aruba.py          (parte col browser VISIBILE)
#
# I SELETTORI dei click vanno riempiti con la "mappa" di Simone (sezioni,
# filtri, bottoni): sono tutti raccolti qui sotto in PASSI, facili da
# correggere a ogni iterazione. In caso di errore lo script salva screenshot
# e HTML della pagina in errori/ — incollali in chat per il giro successivo.
#
# NB: schedulare con l'Utilita' di pianificazione di Windows quando funziona.

import json
import sys
import time
from datetime import datetime
from pathlib import Path

QUI = Path(__file__).parent
CONFIG = QUI / "config.json"
SCARICATI = QUI / "scaricati"
ERRORI = QUI / "errori"

# ---------------------------------------------------------------------------
# LA MAPPA DEI CLICK (da compilare con le indicazioni di Simone).
# Ogni passo: ("descrizione", "selettore css o testo").  I selettori di tipo
# "text=..." cliccano l'elemento che CONTIENE quel testo.
# ---------------------------------------------------------------------------
URL_LOGIN = "https://fatturazioneelettronica.aruba.it/"

PASSI_PRIMA_NOTA: list[tuple[str, str]] = [
    # esempio: ("apri il menu Incassi e pagamenti", "text=Incassi e pagamenti"),
    # esempio: ("premi Esporta", "text=Esporta"),
]

PASSI_NOTE_CREDITO: list[tuple[str, str]] = []


def carica_config() -> dict:
    if not CONFIG.exists():
        esempio = {
            "username": "IL_TUO_UTENTE_ARUBA",
            "password": "LA_TUA_PASSWORD",
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
    page.screenshot(path=str(ERRORI / f"{stamp}-{nome}.png"), full_page=True)
    (ERRORI / f"{stamp}-{nome}.html").write_text(page.content(), encoding="utf-8")
    print(f"ERRORE al passo '{nome}': screenshot e HTML salvati in {ERRORI}")


def esegui_passi(page, passi, etichetta: str) -> None:
    for descrizione, selettore in passi:
        print(f"  → {descrizione}")
        try:
            page.click(selettore, timeout=15000)
            time.sleep(1.5)
        except Exception:
            dump_errore(page, f"{etichetta}-{descrizione[:30]}")
            raise


def main() -> None:
    cfg = carica_config()
    SCARICATI.mkdir(exist_ok=True)
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=bool(cfg.get("headless", False)))
        page = browser.new_page(accept_downloads=True)
        print("Apro Aruba Fatturazione…")
        page.goto(URL_LOGIN)
        # --- LOGIN (selettori da verificare alla prima corsa) ---------------
        try:
            page.fill('input[name="username"], input[type="email"]', cfg["username"])
            page.fill('input[name="password"], input[type="password"]', cfg["password"])
            page.click('button[type="submit"], text=Accedi')
            page.wait_for_load_state("networkidle", timeout=30000)
        except Exception:
            dump_errore(page, "login")
            raise
        print("Login ok (se la pagina e' quella giusta).")

        for etichetta, passi in [
            ("prima-nota", PASSI_PRIMA_NOTA),
            ("note-credito", PASSI_NOTE_CREDITO),
        ]:
            if not passi:
                print(f"[{etichetta}] passi non ancora compilati: salto.")
                continue
            print(f"[{etichetta}] eseguo i passi…")
            with page.expect_download(timeout=60000) as attesa:
                esegui_passi(page, passi, etichetta)
            download = attesa.value
            dest = SCARICATI / f"{datetime.now():%Y%m%d}-{etichetta}-{download.suggested_filename}"
            download.save_as(str(dest))
            print(f"[{etichetta}] scaricato: {dest}")

        browser.close()
    print("Fatto.")


if __name__ == "__main__":
    main()
