@echo off
REM Atura NOMES l'agent (Ollama i Outlook segueixen vius)
echo Aturant l'agent...
for /f "skip=1 tokens=2 delims=," %%p in ('wmic process where "name='python.exe'" get processid^,commandline /format:csv 2^>nul ^| find /I "agente.py"') do (
    taskkill /PID %%p /F /T >nul 2>&1
)
wmic process where "name='python.exe'" get commandline 2>nul | find /I "agente.py" >nul
if errorlevel 1 (echo Agent aturat.) else (echo AVIS: pot quedar algun proces; mira l'Administrador de tareas.)
