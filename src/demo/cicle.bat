@echo off
REM ===================================================================
REM  UN CICLE COMPLET: organitzar la safata i redactar els esborranys.
REM  S'executa UNA vegada i para. No es cap bucle.
REM
REM  Us:   cicle.bat            (350 correus)
REM        cicle.bat 50         (50 correus)
REM        cicle.bat 50 dry     (assaig: no mou ni crea res)
REM
REM  Ho deixa tot escrit a logs\cicle_<data>.txt
REM ===================================================================

setlocal
cd /d C:\Agente_Mail_tal

set MAXM=%1
if "%MAXM%"=="" set MAXM=350
set DRY=
if /i "%2"=="dry" set DRY=--dry

if not exist logs mkdir logs
set F=%date:~-4%%date:~3,2%%date:~0,2%_%time:~0,2%%time:~3,2%
set F=%F: =0%
set LOG=logs\cicle_%F%.txt

echo ================================================= > "%LOG%"
echo  CICLE COMPLET  %date% %time%                    >> "%LOG%"
echo  Maxim: %MAXM%   %DRY%                           >> "%LOG%"
echo ================================================= >> "%LOG%"

echo.
echo ===== 1/4  ORGANITZANT LA SAFATA =====
echo. >> "%LOG%"
echo ---------- ORGANIZAR (%time%) ---------- >> "%LOG%"
python src\demo\organizar.py --outlook "info" --carpeta "Bandeja de entrada" --max %MAXM% %DRY% >> "%LOG%" 2>&1
if errorlevel 1 (
  echo    ATENCIO: l'organizar ha acabat amb error. >> "%LOG%"
  echo    ATENCIO: l'organizar ha acabat amb error, es continua.
)

call :calaix "1 FACIL"        2
call :calaix "2 DIFICIL"      3
call :calaix "3 DESISTIMIENTO" 4

echo. >> "%LOG%"
echo ========== CICLE ACABAT  %date% %time% ========== >> "%LOG%"
echo.
echo Fet. Registre a: %LOG%
echo Revisa els esborranys a la carpeta Esborranys d'Outlook.
echo.
echo Per veure pregunta i resposta de cada correu:
echo    python src\demo\revision_diaria.py
endlocal
exit /b

:calaix
echo.
echo ===== %~2/4  ESBORRANYS: %~1 =====
echo. >> "%LOG%"
echo ---------- %~1 (%time%) ---------- >> "%LOG%"
python src\demo\borradores.py --outlook "info" --carpeta %1 --real --max %MAXM% %DRY% >> "%LOG%" 2>&1
if errorlevel 1 (
  echo    ATENCIO: error al calaix %~1, es continua. >> "%LOG%"
  echo    ATENCIO: error al calaix %~1, es continua.
)
exit /b
