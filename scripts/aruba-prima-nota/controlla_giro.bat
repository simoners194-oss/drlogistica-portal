@echo off
REM DR Portal — GIRO SU RICHIESTA: gira ogni ~10 minuti dall'Utilita' di
REM pianificazione. Chiede al portale se il bottone "Sincronizza da Aruba"
REM ha lasciato una richiesta; se si', esegue il giro completo (prima nota,
REM incassi, collegamenti NC e stati di pagamento). Silenzioso se non c'e'
REM nulla da fare.
cd /d "%~dp0"
set PYEXE=C:\Python314\python.exe
if not exist "%PYEXE%" set PYEXE=python
"%PYEXE%" poll_giro.py
if errorlevel 1 exit /b 0
echo ----- GIRO SU RICHIESTA ----- >> log_esecuzioni.txt
call esegui_giornaliero.bat
