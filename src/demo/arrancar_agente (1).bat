@echo off
REM ================================================================
REM  ARRENCADOR DEL PILOT 24/7 - aixeca Outlook + Agent
REM  Es pot executar mil cops: nomes arrenca el que falti (idempotent)
REM  Pensat per a la tasca programada "en iniciar sessio"
REM ================================================================
cd /d C:\Agente_Mail_Tal

REM --- 1. OUTLOOK: obrir-lo si no corre (el COM el necessita viu) ---
tasklist /FI "IMAGENAME eq OUTLOOK.EXE" 2>nul | find /I "OUTLOOK.EXE" >nul
if errorlevel 1 (
    echo [%time%] Arrencant Outlook...
    start "" outlook.exe
    timeout /t 25 /nobreak >nul
)

REM --- 2. AGENT: nomes si no n'hi ha ja un corrent (anti-duplicats) ---
wmic process where "name='python.exe'" get commandline 2>nul | find /I "agente.py" >nul
if not errorlevel 1 (
    echo [%time%] L'agent ja esta corrent. Res a fer.
    goto :fi
)

REM --- rotacio del log si passa de ~20 MB (l'historial vell es conserva) ---
if exist agente_log.txt (
    for %%A in (agente_log.txt) do if %%~zA GTR 20000000 (
        move /Y agente_log.txt agente_log_anterior.txt >nul
    )
)

echo [%time%] Arrencant l'agent (log: agente_log.txt)...
start "AGENTE" /min cmd /c "python src\demo\agente.py >> agente_log.txt 2>&1"
echo [%time%] Tot en marxa.

:fi
