@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0assets\crear_acceso.ps1"
if errorlevel 1 pause
