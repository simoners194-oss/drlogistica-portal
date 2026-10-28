# DR Portal - installazione delle attivita' pianificate dei giri Aruba su un
# PC nuovo. Si lancia con installa.bat (stessa cartella) da un prompt NORMALE
# dell'utente che terra' acceso il PC: le attivita' girano nella sua sessione,
# con i suoi pacchetti Python.
#
# In ordine:
#   1. cartella degli script (quella sopra a questa) e config.json
#   2. Python 3 e Playwright (installa Playwright se manca)
#   3. Edge o Chrome di sistema (lo scarico usa quelli, invisibili)
#   4. percorso di Python nei due .bat, se non e' C:\Python314
#   5. registrazione delle attivita' "DR Portal - Aruba giornaliero" e
#      "DR Portal - giro su richiesta" per l'utente corrente, senza finestre
#   6. sospensione del PC (facoltativo) e corsa di prova (facoltativa)
# Si puo' rilanciare quante volte si vuole: sovrascrive le attivita'.
# Solo caratteri ASCII in questo file: PowerShell 5.1 legge male gli accenti.

$ErrorActionPreference = "Stop"
$nomeGiornaliero = "DR Portal - Aruba giornaliero"
$nomeRichiesta = "DR Portal - giro su richiesta"

function Titolo($t) { Write-Host ""; Write-Host "== $t" -ForegroundColor Cyan }
function Ok($t) { Write-Host "   OK  $t" -ForegroundColor Green }
function Avviso($t) { Write-Host "   !!  $t" -ForegroundColor Yellow }
function Chiudi($codice) { Write-Host ""; Read-Host "Premi Invio per chiudere" | Out-Null; exit $codice }
function Errore($t) { Write-Host ""; Write-Host "   ERRORE: $t" -ForegroundColor Red; Chiudi 1 }

# Lancia un programma senza finestre e senza che lo stderr diventi un errore
# di PowerShell; restituisce codice di uscita e testo prodotto.
function Esegui($exe, [string[]]$argomenti) {
  $o = [IO.Path]::GetTempFileName(); $e = [IO.Path]::GetTempFileName()
  try {
    $p = Start-Process -FilePath $exe -ArgumentList $argomenti -NoNewWindow -Wait -PassThru -RedirectStandardOutput $o -RedirectStandardError $e
    $testo = [string](Get-Content $o -Raw) + [string](Get-Content $e -Raw)
    return @{ codice = $p.ExitCode; testo = $testo }
  } catch {
    return @{ codice = 9009; testo = $_.Exception.Message }
  } finally {
    Remove-Item $o, $e -ErrorAction SilentlyContinue
  }
}

# Prova un interprete: restituisce il percorso dell'eseguibile se e' un
# Python 3.10+ vero (non l'alias dello Store in WindowsApps), altrimenti $null.
function Prova-Python($exe, [string[]]$pre) {
  $r = Esegui $exe ($pre + @('-c', '"import sys;print(sys.version_info[0],sys.version_info[1],sys.executable)"'))
  if ($r.codice -ne 0) { return $null }
  $righe = @(([string]$r.testo).Trim() -split "`r?`n")
  $parti = $righe[$righe.Count - 1] -split " ", 3
  if ($parti.Count -eq 3 -and $parti[0] -eq "3" -and [int]$parti[1] -ge 10 -and $parti[2] -notlike "*WindowsApps*") { return $parti[2].Trim() }
  return $null
}

function Registra($nome, $template, $cartella, $utente, $inizio) {
  $xml = [IO.File]::ReadAllText((Join-Path $PSScriptRoot $template))
  $xml = $xml.Replace("__CARTELLA__", [Security.SecurityElement]::Escape($cartella))
  $xml = $xml.Replace("__UTENTE__", [Security.SecurityElement]::Escape($utente))
  $xml = $xml.Replace("__INIZIO__", $inizio)
  Register-ScheduledTask -TaskName $nome -Xml $xml -User $utente -Force | Out-Null
  Ok $nome
}

try {
  $cartella = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
  $utente = "$env:USERDOMAIN\$env:USERNAME"

  Titolo "Cartella degli script"
  foreach ($f in "scarica_aruba.py", "poll_giro.py", "esegui_giornaliero.bat", "controlla_giro.bat", "avvia_nascosto.js", "xlrd\__init__.py") {
    if (-not (Test-Path (Join-Path $cartella $f))) { Errore "manca $f in $cartella : il pacchetto e' incompleto" }
  }
  Ok $cartella
  Ok "utente: $utente"

  Titolo "config.json (credenziali Aruba e indirizzo del portale)"
  $cfgPath = Join-Path $cartella "config.json"
  if (-not (Test-Path $cfgPath)) {
    Errore "manca config.json in $cartella`n   Va COPIATO dal PC vecchio (stessa cartella), non ricreato: contiene utente e password Aruba e l'indirizzo del portale con il token. Copialo e rilancia installa.bat."
  }
  $cfg = $null
  try { $cfg = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { Errore "config.json non e' un JSON valido: $($_.Exception.Message)" }
  foreach ($k in "username", "password", "cron_fatture_url") {
    if (-not $cfg.$k) { Errore "in config.json manca il campo '$k'" }
  }
  if ($cfg.cron_fatture_url -notlike "*/cron-fatture*") { Errore "cron_fatture_url in config.json non contiene /cron-fatture" }
  if ($cfg.headless -eq $true) { Ok "config.json completo, headless true" }
  else { Avviso "headless non e' true: il browser dello scarico sarebbe VISIBILE a ogni giro. Metti ""headless"": true in config.json" }

  Titolo "Python"
  $py = $null
  if (Test-Path "C:\Python314\python.exe") { $py = Prova-Python "C:\Python314\python.exe" @() }
  if (-not $py) { $py = Prova-Python "py" @("-3") }
  if (-not $py) { $py = Prova-Python "python" @() }
  if (-not $py) { Errore "Python 3.10 o piu' recente non trovato. Installalo da python.org (Install Now, con la spunta 'Add python.exe to PATH') e rilancia installa.bat." }
  Ok "Python: $py"

  Titolo "Playwright (il pacchetto che guida il browser)"
  function Ha-Playwright { return ((Esegui $py @('-c', '"import playwright"')).codice -eq 0) }
  if (Ha-Playwright) { Ok "Playwright presente" }
  else {
    Avviso "Playwright manca: lo installo adesso (pip install playwright), ci vuole qualche minuto"
    $r = Esegui $py @('-m', 'pip', 'install', '--disable-pip-version-check', 'playwright')
    if (-not (Ha-Playwright)) {
      Write-Host $r.testo
      Errore "Playwright non si e' installato. Da un prompt dei comandi lancia:  `"$py`" -m pip install playwright  e poi rilancia installa.bat"
    }
    Ok "Playwright installato"
  }

  Titolo "Browser di sistema"
  $browser = @(
    "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    "C:\Program Files\Google\Chrome\Application\chrome.exe",
    "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
  ) | Where-Object { Test-Path $_ } | Select-Object -First 1
  if ($browser) { Ok $browser }
  else { Avviso "ne' Edge ne' Chrome trovati: lo scarico userebbe il Chromium di Playwright (da scaricare con  `"$py`" -m playwright install chromium ), che Smart App Control puo' bloccare. Meglio installare Edge." }

  Titolo "Interprete Python nei due .bat"
  if ($py -ieq "C:\Python314\python.exe") { Ok "C:\Python314\python.exe: i .bat vanno bene cosi'" }
  else {
    foreach ($b in "esegui_giornaliero.bat", "controlla_giro.bat") {
      $p = Join-Path $cartella $b
      $txt = [IO.File]::ReadAllText($p)
      $nuovo = [regex]::Replace($txt, "(?m)^set PYEXE=[^\r\n]*", "set PYEXE=$py")
      if ($nuovo -ne $txt) {
        [IO.File]::WriteAllText($p, $nuovo, (New-Object Text.UTF8Encoding $false))
        Ok "$b -> PYEXE=$py"
      } else { Ok "$b gia' a posto" }
    }
  }

  Titolo "Attivita' pianificate (senza finestre: partono tramite wscript)"
  $inizio = (Get-Date).AddMinutes(2).ToString("yyyy-MM-ddTHH:mm:ss")
  Registra $nomeGiornaliero "aruba_giornaliero.template.xml" $cartella $utente $inizio
  Registra $nomeRichiesta "giro_su_richiesta.template.xml" $cartella $utente $inizio
  Write-Host "   Orari: lun-sab 9:05, 12:05, 15:05, 16:35; controllo del giro su richiesta ogni 10 minuti."

  Titolo "Sovrapposizione con il PC vecchio"
  Write-Host "   Le stesse attivita' NON devono girare su due PC insieme (giri in parallelo = errori 429 sul portale)."
  $r = Read-Host "   Sul PC vecchio le due attivita' sono gia' DISATTIVATE? [s/n]"
  if ($r -notmatch "^[sS]") {
    Disable-ScheduledTask -TaskName $nomeGiornaliero | Out-Null
    Disable-ScheduledTask -TaskName $nomeRichiesta | Out-Null
    Avviso "attivita' registrate ma DISATTIVATE. Quando sul PC vecchio sono spente, rilancia installa.bat e rispondi s."
    Chiudi 0
  }
  Ok "attivita' attive"

  Titolo "Sospensione del PC"
  Write-Host "   Le attivita' girano solo con il PC acceso e l'utente $utente connesso (schermo bloccato va bene)."
  $r = Read-Host "   Disattivo la sospensione automatica con alimentazione di rete (powercfg)? [s/n]"
  if ($r -match "^[sS]") {
    $null = Esegui "powercfg" @('/change', 'standby-timeout-ac', '0')
    $null = Esegui "powercfg" @('/change', 'hibernate-timeout-ac', '0')
    Ok "sospensione e ibernazione disattivate con alimentazione di rete"
  }

  Titolo "Corsa di prova"
  $log = Join-Path $cartella "log_esecuzioni.txt"
  $r = Read-Host "   Lancio adesso una corsa completa di prova (5-10 minuti, nessuna finestra)? [s/n]"
  if ($r -match "^[sS]") {
    $prima = 0
    if (Test-Path $log) { $prima = (Get-Item $log).Length }
    Start-ScheduledTask -TaskName $nomeGiornaliero
    Write-Host "   in corso: questa finestra aspetta la fine (massimo 30 minuti)..."
    $t0 = Get-Date
    Start-Sleep -Seconds 8
    while (((Get-ScheduledTask -TaskName $nomeGiornaliero).State -eq "Running") -and (((Get-Date) - $t0).TotalMinutes -lt 30)) { Start-Sleep -Seconds 10 }
    if (Test-Path $log) {
      $fs = [IO.File]::Open($log, 'Open', 'Read', 'ReadWrite')
      $null = $fs.Seek($prima, 'Begin')
      $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::Default)
      $nuovo = $sr.ReadToEnd(); $sr.Close()
      Write-Host ""; Write-Host "   Righe chiave del log:"
      $nuovo -split "`r?`n" | Where-Object { $_ -match "interprete:|Browser avviato|\[stipendi\]|Traceback|Error|ERRORE|Web Page Blocked|fine corsa|cron-incassi|OK:" } | ForEach-Object { Write-Host "     $_" }
      if (($nuovo -match "Browser avviato") -and ($nuovo -match "\[stipendi\] OK") -and ($nuovo -notmatch "Traceback")) { Ok "corsa di prova riuscita" }
      else { Avviso "qualcosa non torna: guarda il log completo $log (e la cartella errori)" }
    } else { Avviso "nessun log scritto: la corsa non e' partita. Utilita' di pianificazione > DR Portal - Aruba giornaliero > Cronologia." }
  }

  Titolo "Fatto"
  Write-Host "   Log di ogni corsa: $log"
  Write-Host "   Per rifare tutto (per esempio dopo aver cambiato Python): doppio clic su installa.bat."
  Chiudi 0
} catch {
  Write-Host ""
  Write-Host "   ERRORE: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "   $($_.ScriptStackTrace)"
  Chiudi 1
}
