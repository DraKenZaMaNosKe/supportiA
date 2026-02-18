@echo off
:: =============================================================================
:: HCG - CORREGIR / INSTALAR TAREAS DE REPORTE
:: =============================================================================
:: Corrige el problema de permisos en tareas existentes
:: Si no existen, las crea desde cero con permisos correctos
:: Se auto-eleva como Administrador
:: =============================================================================

:: --- Auto-elevacion como Administrador ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Solicitando permisos de Administrador...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

title HCG - Corregir Tareas de Reporte
color 0B
cls

echo.
echo   ============================================================
echo   HCG - CORREGIR / INSTALAR TAREAS DE REPORTE
echo   ============================================================
echo   Corrige permisos y recrea tareas programadas
echo   ============================================================
echo.

:: Ejecutar el script PowerShell desde la misma carpeta
if exist "%~dp0InstalarReportes.ps1" (
    powershell -ExecutionPolicy Bypass -File "%~dp0InstalarReportes.ps1"
) else (
    echo   [ERROR] No se encontro InstalarReportes.ps1 en:
    echo   %~dp0
    echo.
    echo   Asegurate de que el archivo este junto a este .bat
    echo.
    pause
)
