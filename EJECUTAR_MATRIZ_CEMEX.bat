@echo off
setlocal enabledelayedexpansion
title CEMEX Agregados - Instalador y Lanzador Automatico
chcp 65001 >nul
cd /d "%~dp0"

echo ===============================================================================
echo                CEMEX AGREGADOS - CONFIGURACION Y LANZADOR
echo ===============================================================================
echo.

:: 1. Buscar Python en PATH y en rutas estandar de Windows
set "PY_EXE="

where python >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=python"
    goto :python_found
)

where py >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=py"
    goto :python_found
)

:: Buscar en carpetas comunes de usuario
for %%V in (313 312 311 310 39) do (
    if exist "%LocalAppData%\Programs\Python\Python%%V\python.exe" (
        set "PY_EXE=%LocalAppData%\Programs\Python\Python%%V\python.exe"
        goto :python_found
    )
    if exist "%ProgramFiles%\Python%%V\python.exe" (
        set "PY_EXE=%ProgramFiles%\Python%%V\python.exe"
        goto :python_found
    )
    if exist "C:\Python%%V\python.exe" (
        set "PY_EXE=C:\Python%%V\python.exe"
        goto :python_found
    )
)

:: Si no se encuentra Python en la maquina, abrir tienda de Microsoft automaticamente
echo [!] Python no esta instalado en este equipo.
echo.
echo [*] Abriendo Microsoft Store para instalar Python con 1 solo clic...
echo     (Solo haz clic en 'Obtener' o 'Instalar' en la ventana que se abre).
echo.
start ms-windows-store://search/?query=Python%%203.11
echo Presiona cualquier tecla una vez que termine de instalarse en la tienda...
pause >nul

:: Reintentar buscar Python
where python >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=python"
    goto :python_found
)

echo [!] Por favor reinicia este archivo una vez completada la instalacion de la tienda.
pause
exit /b 1

:python_found
echo [OK] Python detectado: %PY_EXE%
echo.

:: 2. Instalar librerias necesarias en segundo plano
echo [*] Verificando componentes (pandas, openpyxl)...
"%PY_EXE%" -c "import pandas, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] Instalando librerias necesarias por primera y unica vez...
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
    echo [OK] Acceso directo creado en tu Escritorio: "CEMEX - Matriz de Precios"
)

:: 4. Iniciar la aplicacion de inmediato
echo.
echo [*] Iniciando interfaz grafica CEMEX...
start "" "%TARGET_SCRIPT%"

timeout /t 2 /nobreak >nul
exit
