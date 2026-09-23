@echo off
setlocal enabledelayedexpansion
title CEMEX Agregados - Instalador y Lanzador 100% Automatico
chcp 65001 >nul
cd /d "%~dp0"

echo ===============================================================================
echo                CEMEX AGREGADOS - CONFIGURACION Y LANZADOR
echo ===============================================================================
echo.

set "PY_EXE="

:: 1. Buscar si Python ya esta instalado en PATH o rutas comunes
:buscar_python
where py >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=py"
    goto :test_python
)

where python >nul 2>&1
if %errorlevel% equ 0 (
    python -c "import sys" >nul 2>&1
    if !errorlevel! equ 0 (
        set "PY_EXE=python"
        goto :python_found
    )
)

where python3 >nul 2>&1
if %errorlevel% equ 0 (
    python3 -c "import sys" >nul 2>&1
    if !errorlevel! equ 0 (
        set "PY_EXE=python3"
        goto :python_found
    )
)

for /d %%D in ("%LocalAppData%\Programs\Python\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        goto :python_found
    )
)

for /d %%D in ("%ProgramFiles%\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        goto :python_found
    )
)

for /d %%D in ("%ProgramFiles(x86)%\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        goto :python_found
    )
)

for /d %%D in ("C:\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        goto :python_found
    )
)

for /d %%D in ("%LocalAppData%\Microsoft\WindowsApps\PythonHardwareCompany.Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        goto :python_found
    )
)

:: 2. Si no esta instalado, DESCARGAR E INSTALAR DIRECTAMENTE DESDE CMD (100% AUTOMATICO)
echo [!] Python no esta presente en este equipo.
echo [*] Descargando e instalando Python automaticamente desde CMD...
echo     (No necesitas hacer nada, esto tomara aprox. 20 segundos)...
echo.

:: Intento 1: Winget (Windows Package Manager oficial de Windows 10/11)
where winget >nul 2>&1
if %errorlevel% equ 0 (
    echo [*] Descargando mediante Winget...
    winget install Python.Python.3.11 --silent --accept-package-agreements --accept-source-agreements >nul 2>&1
    if !errorlevel! equ 0 (
        echo [OK] Python instalado exitosamente via Winget.
        goto :buscar_python
    )
)

:: Intento 2: Curl oficial + Instalacion silenciosa de Python oficial
echo [*] Descargando instalador oficial de Python via Curl...
set "INSTALLER_TMP=%TEMP%\python_installer_cemex.exe"
curl -L -s -o "!INSTALLER_TMP!" "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"

if exist "!INSTALLER_TMP!" (
    echo [*] Instalando Python en segundo plano con PATH habilitado...
    "!INSTALLER_TMP!" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 SimpleInstall=1
    del /f /q "!INSTALLER_TMP!" >nul 2>&1
    timeout /t 5 /nobreak >nul
    goto :buscar_python
)

:: Si fallo conexion o bloqueo estricto
echo [!] No se pudo descargar automaticamente por politicas de red.
echo [*] Abriendo la tienda Microsoft Store para 1 clic manual...
start ms-windows-store://search/?query=Python
pause
exit /b 1

:test_python
"%PY_EXE%" -c "import sys; print(sys.version)" >nul 2>&1
if %errorlevel% neq 0 (
    set "PY_EXE="
    goto :buscar_python
)

:python_found
echo [OK] Python detectado correctamente: %PY_EXE%
"%PY_EXE%" -c "import sys; print('     Version: ' + sys.version.split()[0])"
echo.

:: 3. Instalar dependencias necesarias automaticamente
echo [*] Verificando componentes (pandas, openpyxl)...
"%PY_EXE%" -c "import pandas, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] Instalando librerias necesarias por primera vez...
    echo     Esto tomara solo unos segundos, por favor espera...
    echo.
    "%PY_EXE%" -m pip install --quiet pandas openpyxl
    if %errorlevel% neq 0 (
        "%PY_EXE%" -m pip install --user --quiet pandas openpyxl
    )
    echo [OK] Componentes listos e instalados con exito.
) else (
    echo [OK] Componentes listos.
)

:: 4. Crear Acceso Directo automatico en el Escritorio
set "SCRIPT_DIR=%~dp0"
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\CEMEX - Matriz de Precios.lnk"
set "TARGET_SCRIPT=%SCRIPT_DIR%Lanzador_CEMEX.pyw"

powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%SHORTCUT_PATH%'); $s.TargetPath = '%TARGET_SCRIPT%'; $s.WorkingDirectory = '%SCRIPT_DIR%'; $s.Description = 'CEMEX Agregados - Matriz de Precios y Costos'; $s.Save()" >nul 2>&1

if exist "%SHORTCUT_PATH%" (
    echo [OK] Acceso directo disponible en tu Escritorio: "CEMEX - Matriz de Precios"
)

:: 5. Iniciar la aplicacion de inmediato
echo.
echo [*] Iniciando interfaz grafica CEMEX...
start "" "%TARGET_SCRIPT%"

timeout /t 2 /nobreak >nul
exit
