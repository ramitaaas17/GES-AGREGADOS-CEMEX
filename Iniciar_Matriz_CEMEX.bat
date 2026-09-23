@echo off
setlocal enabledelayedexpansion
title CEMEX Agregados - Sistema Integral de Matriz y Gobernanza
chcp 65001 >nul
cd /d "%~dp0"

echo ===============================================================================
echo                CEMEX AGREGADOS - SISTEMA INTEGRAL DE MATRIZ
echo ===============================================================================
echo.

:: 1. Verificar si Python está instalado en el sistema
where python >nul 2>&1
if %errorlevel% neq 0 (
    where py >nul 2>&1
    if %errorlevel% neq 0 (
        echo [!] ERROR: Python no se encuentra instalado en este equipo.
        echo.
        echo Para usar el sistema, por favor instale Python desde:
        echo https://www.python.org/downloads/ o desde Microsoft Store.
        echo.
        echo (Asegúrese de marcar la casilla "Add Python to PATH" durante la instalación).
        echo.
        pause
        exit /b 1
    ) else (
        set PY_CMD=py
    )
) else (
    set PY_CMD=python
)

:: 2. Verificar e instalar librerías requeridas (pandas, openpyxl)
echo [*] Verificando componentes de Python (pandas, openpyxl)...
%PY_CMD% -c "import pandas, openpyxl; print('OK')" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] Instalando librerías necesarias por primera vez...
    echo     Esto solo tomará unos segundos.
    echo.
    %PY_CMD% -m pip install --upgrade pip >nul 2>&1
    %PY_CMD% -m pip install pandas openpyxl
    if %errorlevel% neq 0 (
        echo.
        echo [*] Intentando instalación en modo usuario...
        %PY_CMD% -m pip install --user pandas openpyxl
    )
    echo.
    echo [OK] Componentes listos e instalados correctamente.
)

:: 3. Iniciar la Aplicación CEMEX
echo.
echo [*] Iniciando interfaz gráfica CEMEX...
start "" "%~dp0\Lanzador_CEMEX.pyw"

timeout /t 2 /nobreak >nul
exit
