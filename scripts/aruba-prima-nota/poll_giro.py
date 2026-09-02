# DR Portal — POLL del giro su richiesta.
# Interroga /cron-giro del portale: se il bottone "Sincronizza da Aruba" ha
# lasciato una richiesta fresca, il server la consuma e risponde ESEGUI.
# Codici di uscita: 0 = eseguire il giro completo, 1 = niente da fare.
# Il lock giro_in_corso.lock evita giri sovrapposti (stale dopo 2 ore).
import json
import sys
import time
import urllib.request
from pathlib import Path

BASE = Path(__file__).resolve().parent
LOCK = BASE / "giro_in_corso.lock"

if LOCK.exists():
    eta = time.time() - LOCK.stat().st_mtime
    if eta < 2 * 3600:
        sys.exit(1)  # giro gia' in corso: si riprova al prossimo poll
    LOCK.unlink(missing_ok=True)  # lock stantio (giro morto male): via

try:
    cfg = json.loads((BASE / "config.json").read_text(encoding="utf-8"))
except Exception:
    sys.exit(1)
base = str(cfg.get("cron_fatture_url", "")).strip()
if not base or "/cron-fatture" not in base:
    sys.exit(1)
url = base.replace("/cron-fatture", "/cron-giro")
req = urllib.request.Request(
    url,
    headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) DRPortalCron/1.0"},
)
try:
    with urllib.request.urlopen(req, timeout=60) as r:
        corpo = r.read().decode("utf-8", "replace").replace("<!-- -->", "")
except Exception:
    sys.exit(1)
sys.exit(0 if "OK: ESEGUI" in corpo else 1)
