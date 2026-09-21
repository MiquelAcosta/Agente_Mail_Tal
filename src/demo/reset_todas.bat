@echo off
REM ===================================================================
REM  REINICIAR PROVES a tots els calaixos: treu l'etiqueta 'Agente' i
REM  esborra la fila del registre, perque els mails es puguin tornar a
REM  processar. NO esborra cap correu.
REM
REM  Us:   reset_todas.bat          -> ensenya que faria (segur)
REM        reset_todas.bat aplicar  -> ho aplica de veritat
REM ===================================================================
setlocal
cd /d C:\Agente_Mail_tal

set MODE=--dry
if /i "%1"=="aplicar" set MODE=

for %%C in ("1 FACIL" "2 DIFICIL" "3 DESISTIMIENTO") do (
  echo.
  echo ===== %%~C =====
  python src\demo\reset_pruebas.py --outlook "info" --carpeta %%C %MODE%
)

echo.
if "%MODE%"=="--dry" (
  echo MODE PROVA: no s'ha canviat res. Executa:  reset_todas.bat aplicar
) else (
  echo Fet. Ja pots llancar:  tanda_completa.bat 10
)
endlocal
