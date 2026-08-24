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
echo ================= %date% %time% ================= >> log_esecuzioni.txt
"%PYEXE%" scarica_aruba.py >> log_esecuzioni.txt 2>&1
"%PYEXE%" scarica_aruba.py incassi >> log_esecuzioni.txt 2>&1
"%PYEXE%" scarica_aruba.py nclinks >> log_esecuzioni.txt 2>&1
echo (fine corsa) >> log_esecuzioni.txt
