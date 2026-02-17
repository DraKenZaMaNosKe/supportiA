@echo off
title HCG - Configurador de Equipos
color 0B
cls
echo.
echo  ============================================================
echo       HOSPITAL CIVIL DE GUADALAJARA
echo       Coordinacion General de Informatica
echo  ============================================================
echo.
echo       Configurador de Equipos HCG
echo.
echo  ============================================================
echo.
echo  Configurando permisos de ejecucion...
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force" >nul 2>&1
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force" >nul 2>&1
echo  [OK] Permisos configurados
echo.
echo  Iniciando configurador...
echo.
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0ConfigurarEquipoHCG.ps1"
echo.
echo  ============================================================
echo  Proceso finalizado. Presione una tecla para cerrar.
echo  ============================================================
pause >nul
