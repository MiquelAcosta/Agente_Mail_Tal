@echo off
REM ===================================================================
REM  TANDA COMPLETA — genera esborranys als tres calaixos, un darrere
REM  l'altre, sense supervisio.
REM
REM  Us:   tanda_completa.bat          (10 mails per calaix)
REM        tanda_completa.bat 25       (25 mails per calaix)
REM
REM  NO toca 4 ESCALADOS: aquells mails porten arxius i no han de
REM  rebre mai esborrany.
REM  Ho deixa tot escrit a logs\tanda_<data>.txt
REM ===================================================================

setlocal
cd /d C:\Agente_Mail_tal

set MAXM=%1
if "%MAXM%"=="" set MAXM=10

if not exist logs mkdir logs
set FECHA=%date:~-4%%date:~3,2%%date:~0,2%_%time:~0,2%%time:~3,2%
set FECHA=%FECHA: =0%
set LOG=logs\tanda_%FECHA%.txt

echo ================================================= > "%LOG%"
echo  TANDA COMPLETA  %date% %time%                   >> "%LOG%"
echo  Maxim per calaix: %MAXM%                        >> "%LOG%"
echo ================================================= >> "%LOG%"

call :calaix "1 FACIL"
call :calaix "2 DIFICIL"
call :calaix "3 DESISTIMIENTO"

echo. >> "%LOG%"
echo ========== TANDA ACABADA  %date% %time% ========== >> "%LOG%"
echo.
echo Fet. Registre a: %LOG%
echo Els esborranys son a la carpeta Esborranys d'Outlook.
endlocal
exit /b

:calaix
echo.
echo =================================================
echo  CALAIX %~1
echo =================================================
echo. >> "%LOG%"
echo ---------- CALAIX %~1  (%time%) ---------- >> "%LOG%"
python src\demo\borradores.py --outlook "info" --carpeta %1 --real --max %MAXM% >> "%LOG%" 2>&1
if errorlevel 1 (
  echo    ATENCIO: el calaix %~1 ha acabat amb error, es continua. >> "%LOG%"
  echo    ATENCIO: error al calaix %~1, es continua amb el seguent.
)
exit /b
