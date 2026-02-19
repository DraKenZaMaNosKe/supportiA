@echo off
:: =============================================================================
:: HCG - CONFIGURAR OOBE AUTOMATICO
:: =============================================================================
:: Ejecutar desde Shift+F10 durante el OOBE de Windows 11
:: Configura todo automatico y reinicia al paso de creacion de usuario
:: =============================================================================

title HCG - Configurando OOBE...
color 0B

echo.
echo   ============================================================
echo   HCG - CONFIGURAR OOBE AUTOMATICO
echo   ============================================================
echo.
echo   Esto va a:
echo     1. Saltar el requisito de red (BypassNRO)
echo     2. Configurar Mexico, teclado latinoamericano
echo     3. Saltar EULA, privacidad, cuenta Microsoft
echo     4. Reiniciar al paso de creacion de usuario
echo.
echo   ============================================================
echo.

:: --- Detectar la letra de la USB donde esta este script ---
set "SCRIPTDRIVE=%~d0"
echo   USB detectada en: %SCRIPTDRIVE%

:: --- 1. Bypass NRO (saltar requisito de red) ---
echo.
echo   [1/4] Aplicando BypassNRO...
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" /v BypassNRO /t REG_DWORD /d 1 /f >nul 2>&1
if %errorlevel% equ 0 (
    echo         OK - BypassNRO aplicado
) else (
    echo         ERROR - No se pudo aplicar BypassNRO
)

:: --- 2. Copiar unattend.xml a la ubicacion correcta ---
echo.
echo   [2/4] Copiando configuracion OOBE...
if not exist "C:\Windows\Panther\Unattend" (
    mkdir "C:\Windows\Panther\Unattend" >nul 2>&1
)

if exist "%SCRIPTDRIVE%\autounattend.xml" (
    copy /y "%SCRIPTDRIVE%\autounattend.xml" "C:\Windows\Panther\Unattend\unattend.xml" >nul 2>&1
    echo         OK - unattend.xml copiado
) else if exist "%SCRIPTDRIVE%\OOBE\autounattend.xml" (
    copy /y "%SCRIPTDRIVE%\OOBE\autounattend.xml" "C:\Windows\Panther\Unattend\unattend.xml" >nul 2>&1
    echo         OK - unattend.xml copiado (desde carpeta OOBE)
) else (
    echo         WARN - No se encontro autounattend.xml, solo se aplica BypassNRO
)

:: --- 3. Configurar region Mexico por registro ---
echo.
echo   [3/4] Configurando region Mexico...
reg add "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation" /v TimeZoneKeyName /t REG_SZ /d "Central Standard Time (Mexico)" /f >nul 2>&1
echo         OK - Zona horaria Mexico Central

:: --- 4. Reiniciar OOBE ---
echo.
echo   [4/4] Reiniciando OOBE...
echo.
echo   ============================================================
echo   La computadora se va a reiniciar en 5 segundos.
echo   Al volver, el OOBE saltara directo a crear usuario.
echo   ============================================================
echo.
timeout /t 5
shutdown /r /t 0
