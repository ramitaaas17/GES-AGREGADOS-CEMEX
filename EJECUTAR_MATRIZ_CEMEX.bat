@echo off
setlocal enabledelayedexpansion
title CEMEX Agregados - Instalador y Lanzador Universal
chcp 65001 >nul
cd /d "%~dp0"

echo ===============================================================================
echo                CEMEX AGREGADOS - CONFIGURACION Y LANZADOR
echo ===============================================================================
echo.

set "PY_EXE="

:: 1. Probar comandos estandar en PATH
where py >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=py"
    goto :test_python
)

where python >nul 2>&1
if %errorlevel% equ 0 (
    :: Verificar que no sea el alias vacio de la tienda
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

:: 2. Busqueda dinamica en carpetas de instalacion de TODAS las versiones (3.8 a 3.15)
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

for /d %%D in ("%ProgramData%\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        goto :python_found
    )
)

:: 3. Busqueda en WindowsApps (Microsoft Store)
for /d %%D in ("%LocalAppData%\Microsoft\WindowsApps\PythonHardwareCompany.Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        goto :python_found
    )
)

:: Si no se encontro ninguna version instalada
echo [!] No se detecto ninguna instalacion de Python en este equipo.
echo.
echo [*] Abriendo Microsoft Store para instalar Python con 1 solo clic...
echo     (Haz clic en 'Obtener' o 'Instalar' en la ventana de la tienda).
echo.
start ms-windows-store://search/?query=Python
echo Presiona cualquier tecla una vez que termine de instalarse en la tienda...
pause >nul

:: Reintentar busqueda general
for /d %%D in ("%LocalAppData%\Programs\Python\Python*") do (
    if exist "%%D\python.exe" (
        set "PY_EXE=%%D\python.exe"
        goto :python_found
    )
)

where python >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=python"
    goto :python_found
)

echo [!] Por favor vuelve a abrir este archivo una vez completada la instalacion.
pause
exit /b 1

:test_python
"%PY_EXE%" -c "import sys; print(sys.version)" >nul 2>&1
if %errorlevel% neq 0 (
    set "PY_EXE="
    goto :python_found
)

:python_found
echo [OK] Python detectado correctamente: %PY_EXE%
"%PY_EXE%" -c "import sys; print('     Version: ' + sys.version.split()[0])"
echo.

:: 2. Instalar dependencias necesarias automaticamente
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

:: 3. Crear Acceso Directo automatico en el Escritorio
set "SCRIPT_DIR=%~dp0"
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\CEMEX - Matriz de Precios.lnk"
set "TARGET_SCRIPT=%SCRIPT_DIR%Lanzador_CEMEX.pyw"

powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%SHORTCUT_PATH%'); $s.TargetPath = '%TARGET_SCRIPT%'; $s.WorkingDirectory = '%SCRIPT_DIR%'; $s.Description = 'CEMEX Agregados - Matriz de Precios y Costos'; $s.Save()" >nul 2>&1

if exist "%SHORTCUT_PATH%" (
    echo [OK] Acceso directo disponible en tu Escritorio: "CEMEX - Matriz de Precios"
)

:: 4. Iniciar la aplicacion de inmediato
echo.
echo [*] Iniciando interfaz grafica CEMEX...
start "" "%TARGET_SCRIPT%"

timeout /t 2 /nobreak >nul
exit
