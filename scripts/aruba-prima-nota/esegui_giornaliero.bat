@echo off
REM DR Portal — scarico giornaliero da Aruba: prima nota (2 anni) e
REM collegamenti NC. Pensato per l'Utilita' di pianificazione di Windows:
REM   Crea attivita' di base -> Giornaliera (es. 7:15) -> Avvia programma
REM   -> questo file. Consiglio: in config.json mettere "headless": true
REM   cosi' il browser lavora invisibile.
REM Il log di ogni corsa finisce in log_esecuzioni.txt qui accanto.
cd /d "%~dp0"
REM Interprete ESPLICITO: sotto l'Utilita' di pianificazione "python" puo'
REM risolvere un interprete diverso (senza xlrd) e il passo incassi moriva
REM in silenzio nel log. Se Python viene aggiornato, aggiornare il percorso.
set PYEXE=C:\Python314\python.exe
if not exist "%PYEXE%" set PYEXE=python
REM Lock anti-sovrapposizione: il giro su richiesta (controlla_giro.bat)
REM aspetta se un giro e' gia' in corso.
echo su > giro_in_corso.lock
echo ================= %date% %time% ================= >> log_esecuzioni.txt
REM Diagnostica: quale interprete gira davvero sotto lo scheduler.
"%PYEXE%" -c "import sys; print('interprete:', sys.executable)" >> log_esecuzioni.txt 2>&1
"%PYEXE%" scarica_aruba.py >> log_esecuzioni.txt 2>&1
"%PYEXE%" scarica_aruba.py incassi >> log_esecuzioni.txt 2>&1
"%PYEXE%" scarica_aruba.py nclinks >> log_esecuzioni.txt 2>&1
REM Stipendi (1.79.0): il portale legge da solo Stipendi Dr.xlsx e i COSTI
REM mensili da SharePoint. Stesso token del cron fatture, cambia solo il path.
REM NB: Cloudflare risponde 403 allo User-Agent di python-urllib (visto il
REM 28/09: 6 corse fallite dal 21/09): serve un UA esplicito come negli altri script.
"%PYEXE%" -c "import json,re,urllib.request;u=json.load(open('config.json'))['cron_fatture_url'].replace('/cron-fatture','/cron-stipendi');r=urllib.request.urlopen(urllib.request.Request(u,headers={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) DRPortalCron/1.0'}),timeout=600).read().decode('utf-8','ignore');r=re.sub(r'<!--.*?-->','',r);m=re.search(r'(OK|ERRORE)[^<]{0,300}',r);print('[stipendi]',m.group(0) if m else r[:200])" >> log_esecuzioni.txt 2>&1
echo (fine corsa) >> log_esecuzioni.txt
del giro_in_corso.lock 2>nul
