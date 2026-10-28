@echo off
REM DR Portal - installa su QUESTO PC le due attivita' pianificate dei giri
REM Aruba ("DR Portal - Aruba giornaliero" e "DR Portal - giro su richiesta").
REM Doppio clic da utente normale (NON "come amministratore": le attivita'
REM girano nella sessione dell'utente che lancia questo file, con i suoi
REM pacchetti Python). Tutto il lavoro lo fa installa.ps1 qui accanto.
REM Si puo' rilanciare quante volte si vuole.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installa.ps1"
