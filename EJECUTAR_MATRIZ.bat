@echo off
setlocal EnableExtensions
chcp 65001 >nul
title CEMEX - Generador de Matriz de Precios

rem ============================================================================
rem  Lanzador de la Matriz de Precios (Python) SIN pasar por Excel.
rem
rem  Se usa en equipos donde la politica de seguridad de la organizacion
rem  (Defender ASR: "Bloquear que las apps de Office creen procesos
rem  secundarios") impide que Excel ejecute programas. Flujo:
rem    1) En Excel se pulsa el boton y se eligen los CEDIS: Excel guarda la
rem       seleccion en %LOCALAPPDATA%\CEMEX_MatrizVentas\solicitud.txt
rem       (y una copia de los datos actuales del libro en datos_fuente.xlsm)
rem    2) Doble clic en este archivo: ejecuta Python y abre el reporte.
rem  Tambien funciona solo: si no hay seleccion guardada, pregunta los CEDIS y usa el
rem  libro CEMEX_MATRIZ_INTEGRADO.xlsm guardado junto a este archivo (guarde antes con
rem  Ctrl+G si acaba de actualizar datos).
rem ============================================================================

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PYTHONDONTWRITEBYTECODE=1"
set "TRABAJO=%LOCALAPPDATA%\CEMEX_MatrizVentas"
set "SALIDA=%TRABAJO%\salidas"
set "CACHE=%TRABAJO%\cache"
set "SOLICITUD=%TRABAJO%\solicitud.txt"
set "SCRIPT=%~dp0matriz_integrada.py"
set "DESTDIR=%~dp0_salidas_integradas"

if not exist "%SCRIPT%" goto :SIN_SCRIPT
if not exist "%SALIDA%" mkdir "%SALIDA%"
if not exist "%CACHE%" mkdir "%CACHE%"

rem --- CEDIS a procesar ---
set "CEDIS="
set "HAY_SOLICITUD="
if exist "%SOLICITUD%" set "HAY_SOLICITUD=1"
if exist "%SOLICITUD%" set /p CEDIS=<"%SOLICITUD%"
if defined CEDIS goto :CEDIS_OK
echo.
echo Escriba los codigos de CEDIS separados por espacio, por ejemplo: D836 D838
echo o escriba TODOS para procesar todos los centros.
set /p CEDIS=CEDIS:
if not defined CEDIS goto :SIN_CEDIS
:CEDIS_OK
set "ARG_CEDIS=--cedis %CEDIS%"
if /i "%CEDIS%"=="TODOS" set "ARG_CEDIS="

rem --- Origen de los datos ---
rem  1) copia que Excel acaba de tomar de los datos actuales (si hay solicitud de Excel)
rem  2) libro integrado guardado junto a este archivo
rem  3) archivo de extraccion antiguo (*2026*.xlsm)
set "COPIA=%TRABAJO%\datos_fuente.xlsm"
set "FUENTE=%~dp0*2026*.xlsm"
if exist "%~dp0CEMEX_MATRIZ_INTEGRADO.xlsm" set "FUENTE=%~dp0CEMEX_MATRIZ_INTEGRADO.xlsm"
if defined HAY_SOLICITUD if exist "%COPIA%" set "FUENTE=%COPIA%"

echo.
echo  CEMEX - Matriz de Precios
echo  CEDIS: %CEDIS%
echo  Datos: %FUENTE%
echo.

rem --- Buscar Python 3 con pandas y openpyxl ---
echo Buscando Python...
set "PY="
call :PROBAR "py -3"
if not defined PY call :PROBAR "python"
if not defined PY for /d %%D in ("%LOCALAPPDATA%\Python\python*" "%LOCALAPPDATA%\Programs\Python\Python*" "%ProgramW6432%\Python*" "%ProgramFiles%\Python*" "C:\Python*") do if not defined PY call :PROBAR_EXE "%%~D\python.exe"
if defined PY goto :PY_OK

set "PYBASE="
call :EXISTE "py -3"
if not defined PYBASE call :EXISTE "python"
if not defined PYBASE goto :SIN_PYTHON
echo Se encontro Python pero faltan pandas / openpyxl. Instalando ^(requiere Internet^)...
%PYBASE% -m pip install --user --disable-pip-version-check pandas openpyxl xlrd
call :PROBAR "%PYBASE%"
if not defined PY goto :SIN_LIBS

:PY_OK
echo Generando matriz, puede tardar de 20 segundos a varios minutos...
echo.
%PY% "%SCRIPT%" --fuente "%FUENTE%" %ARG_CEDIS% --output "%SALIDA%" --cache-dir "%CACHE%"
set "RC=%errorlevel%"
del "%SOLICITUD%" >nul 2>&1
if not "%RC%"=="0" goto :FALLO

rem --- Reporte mas reciente ---
set "RESULT="
for /f "delims=" %%F in ('dir /b /a-d /o-d "%SALIDA%\Matriz_Precios_Integral_*.xlsx" 2^>nul') do if not defined RESULT set "RESULT=%SALIDA%\%%F"
if not defined RESULT goto :FALLO

rem --- Copiar a la carpeta del proyecto (lo hace PowerShell; cmd/python no pueden escribir en Documentos con Acceso controlado a carpetas) ---
set "FINAL=%RESULT%"
set "SRC=%RESULT%"
for %%F in ("%RESULT%") do set "NOMBRE=%%~nxF"
powershell -NoProfile -Command "New-Item -ItemType Directory -Force -Path $env:DESTDIR | Out-Null; Copy-Item -LiteralPath $env:SRC -Destination $env:DESTDIR -Force" >nul 2>&1
if exist "%DESTDIR%\%NOMBRE%" set "FINAL=%DESTDIR%\%NOMBRE%"

echo.
echo  Listo. Reporte generado:
echo  %FINAL%
echo.
if not defined CEMEX_NO_ABRIR start "" "%FINAL%"
goto :FIN

:FALLO
echo.
echo  No se pudo generar la matriz. Revise los mensajes de arriba ^(codigo %RC%^).
goto :FIN

:SIN_SCRIPT
echo No se encontro matriz_integrada.py junto a este archivo.
goto :FIN

:SIN_CEDIS
echo No se indico ningun CEDIS.
if exist "%SOLICITUD%" del "%SOLICITUD%" >nul 2>&1
goto :FIN

:SIN_PYTHON
echo No se encontro Python 3 en este equipo.
echo Instalelo desde el Centro de Software o https://www.python.org/downloads/
echo ^(marque "Add python.exe to PATH"^) y vuelva a ejecutar este archivo.
goto :FIN

:SIN_LIBS
echo No se pudieron instalar pandas / openpyxl. Verifique la conexion a Internet.
goto :FIN

:FIN
echo.
if not defined CEMEX_SIN_PAUSA pause
endlocal
exit /b

:PROBAR
%~1 -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 exit /b 1
set "PY=%~1"
exit /b 0

:PROBAR_EXE
if not exist "%~1" exit /b 1
"%~1" -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 exit /b 1
set PY="%~1"
exit /b 0

:EXISTE
%~1 -c "import sys" >nul 2>&1
if not errorlevel 1 set "PYBASE=%~1"
exit /b 0
