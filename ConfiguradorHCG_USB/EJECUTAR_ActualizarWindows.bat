@echo off
title HCG - Actualizador de Windows
color 0B
cls
echo.
echo  ============================================================
echo       HOSPITAL CIVIL DE GUADALAJARA
echo       Actualizador de Windows
echo  ============================================================
echo.
echo  Configurando permisos de ejecucion...
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force" >nul 2>&1
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force" >nul 2>&1
echo  [OK] Permisos configurados
echo.
echo  Iniciando actualizador...
echo.
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0ActualizarWindows.ps1"
echo.
echo  ============================================================
echo  Proceso finalizado. Presione una tecla para cerrar.
echo  ============================================================
pause >nul
