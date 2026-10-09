@echo off
setlocal enabledelayedexpansion
title CEMEX Agregados - Lanzador del Sistema
chcp 65001 >nul
cd /d "%~dp0"

echo ===============================================================================
echo                CEMEX AGREGADOS - CONFIGURACION Y LANZADOR
echo ===============================================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "TARGET_SCRIPT=%SCRIPT_DIR%Lanzador_CEMEX.pyw"

:: 1. Buscar Python en el equipo
set "PY_EXE="
set "PYW_EXE="

echo [*] Verificando instalacion de Python en este equipo...

where python >nul 2>&1
if %errorlevel% equ 0 (
    python -c "import sys" >nul 2>&1
    if !errorlevel! equ 0 (
        set "PY_EXE=python"
        where pythonw >nul 2>&1
        if !errorlevel! equ 0 set "PYW_EXE=pythonw"
        goto :python_found
    )
)

where py >nul 2>&1
if %errorlevel% equ 0 (
    py -3 -c "import sys" >nul 2>&1
    if !errorlevel! equ 0 (
        set "PY_EXE=py -3"
        set "PYW_EXE=pyw -3"
        goto :python_found
    )
)

for /d %%D in ("%LocalAppData%\Programs\Python\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        if exist "%%D\pythonw.exe" set "PYW_EXE=%%D\pythonw.exe"
        goto :python_found
    )
)

for /d %%D in ("%ProgramFiles%\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        if exist "%%D\pythonw.exe" set "PYW_EXE=%%D\pythonw.exe"
        goto :python_found
    )
)

for /d %%D in ("%ProgramFiles(x86)%\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        if exist "%%D\pythonw.exe" set "PYW_EXE=%%D\pythonw.exe"
        goto :python_found
    )
)

for /d %%D in ("C:\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        if exist "%%D\pythonw.exe" set "PYW_EXE=%%D\pythonw.exe"
        goto :python_found
    )
)

for /d %%D in ("%ProgramData%\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        if exist "%%D\pythonw.exe" set "PYW_EXE=%%D\pythonw.exe"
        goto :python_found
    )
)

:: Si no se encontro Python
echo [!] No se encontro Python en este equipo.
echo [*] Abriendo Microsoft Store para instalar Python...
start ms-windows-store://search/?query=Python
echo.
echo Presiona cualquier tecla una vez instalado para continuar...
pause >nul
exit /b 1

:python_found
echo [OK] Python detectado: !PY_EXE!
echo.

:: 2. Verificar dependencias requeridas (pandas, openpyxl)
echo [*] Verificando componentes: pandas y openpyxl...
!PY_EXE! -c "import pandas, openpyxl" >nul 2>&1
if !errorlevel! neq 0 (
    echo [*] Instalando componentes por primera vez: pandas y openpyxl...
    !PY_EXE! -m pip install pandas openpyxl
    if !errorlevel! neq 0 (
        !PY_EXE! -m pip install --user pandas openpyxl
    )
    echo [OK] Componentes listos.
) else (
    echo [OK] Componentes listos.
)

:: 3. Crear Acceso Directo automatico en el Escritorio
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\CEMEX - Matriz de Precios.lnk"
if not exist "%SHORTCUT_PATH%" (
    powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%SHORTCUT_PATH%'); $s.TargetPath = '%TARGET_SCRIPT%'; $s.WorkingDirectory = '%SCRIPT_DIR%'; $s.Description = 'CEMEX Agregados - Matriz de Precios y Costos'; $s.Save()" >nul 2>&1
    if exist "%SHORTCUT_PATH%" (
        echo [OK] Acceso directo disponible en tu Escritorio: "CEMEX - Matriz de Precios"
    )
)

:: 4. Iniciar la aplicacion de inmediato
echo.
echo [*] Iniciando interfaz grafica CEMEX...
if defined PYW_EXE (
    start "" !PYW_EXE! "%TARGET_SCRIPT%"
) else (
    start "" !PY_EXE! "%TARGET_SCRIPT%"
)

exit /b 0
