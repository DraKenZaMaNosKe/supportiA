@echo off
title Cosmos Upgrade - Windows 11
color 0B
echo.
echo   Iniciando Cosmos Upgrade - Windows 11...
echo.
powershell.exe -ExecutionPolicy Bypass -File "%~dp0ActualizarWindows11.ps1"
pause
