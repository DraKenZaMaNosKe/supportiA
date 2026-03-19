@echo off
:: =============================================================================
:: HCG - CONFIGURAR OOBE AUTOMATICO (v3)
:: =============================================================================
:: Ejecutar desde Shift+F10 durante el OOBE de Windows 11
:: Configura todo automatico y reinicia al paso de creacion de usuario
:: =============================================================================

title HCG - Configurando OOBE...
color 0B

echo.
echo   ============================================================
echo   HCG - CONFIGURAR OOBE AUTOMATICO v3
echo   ============================================================
echo.
echo   Esto va a:
echo     1. Saltar requisito de red (BypassNRO)
echo     2. Saltar pantallas de privacidad (todo activado)
echo     3. Configurar Mexico, teclado latinoamericano
echo     4. Saltar segundo teclado, EULA, proteccion
echo     5. Copiar unattend.xml
echo     6. Reiniciar al paso de creacion de usuario
echo.
echo   ============================================================
echo.

:: --- Detectar la letra de la USB donde esta este script ---
set "SCRIPTDRIVE=%~d0"
set "SCRIPTPATH=%~dp0"
echo   USB detectada en: %SCRIPTDRIVE%
echo   Ruta del script: %SCRIPTPATH%

:: --- 1. Bypass NRO (saltar requisito de red) ---
echo.
echo   [1/6] Aplicando BypassNRO...
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" /v BypassNRO /t REG_DWORD /d 1 /f >nul 2>&1
if %errorlevel% equ 0 (
    echo         OK - BypassNRO aplicado
) else (
    echo         ERROR - No se pudo aplicar BypassNRO
)

:: --- 2. Saltar pantallas de privacidad y pre-activar todo ---
echo.
echo   [2/6] Configurando privacidad (activar todo y saltar pantalla)...
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\OOBE" /v DisablePrivacyExperience /t REG_DWORD /d 1 /f >nul 2>&1
echo         OK - Pantalla de privacidad deshabilitada

:: Ubicacion - Activar
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location" /v Value /t REG_SZ /d "Allow" /f >nul 2>&1
echo         OK - Ubicacion activada

:: Encontrar mi dispositivo - Activar
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\userNotificationListener" /v Value /t REG_SZ /d "Allow" /f >nul 2>&1
echo         OK - Encontrar mi dispositivo activado

:: Datos de diagnostico - Activar (nivel completo)
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" /v AllowTelemetry /t REG_DWORD /d 3 /f >nul 2>&1
echo         OK - Datos de diagnostico activados

:: Experiencias personalizadas - Activar
reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Privacy" /v TailoredExperiencesWithDiagnosticDataEnabled /t REG_DWORD /d 1 /f >nul 2>&1
echo         OK - Experiencias personalizadas activadas

:: Mejora de entrada manuscrita - Activar
reg add "HKCU\SOFTWARE\Microsoft\Input\TIPC" /v Enabled /t REG_DWORD /d 1 /f >nul 2>&1
echo         OK - Mejora de entrada activada

:: Publicidad personalizada - Activar
reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo" /v Enabled /t REG_DWORD /d 1 /f >nul 2>&1
echo         OK - Publicidad personalizada activada

:: --- 3. Configurar region e idioma Mexico por registro ---
echo.
echo   [3/6] Configurando region Mexico...
reg add "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation" /v TimeZoneKeyName /t REG_SZ /d "Central Standard Time (Mexico)" /f >nul 2>&1
echo         OK - Zona horaria Mexico Central

:: Pre-configurar teclado latinoamericano y saltar segundo teclado
reg add "HKCU\Keyboard Layout\Preload" /v 1 /t REG_SZ /d "0000080a" /f >nul 2>&1
echo         OK - Teclado latinoamericano pre-configurado

:: Eliminar cualquier segundo teclado pre-cargado
reg delete "HKCU\Keyboard Layout\Preload" /v 2 /f >nul 2>&1
echo         OK - Segundo teclado eliminado

:: --- 4. Saltar EULA y pantalla de proteccion ---
echo.
echo   [4/6] Configurando EULA y proteccion...

:: Marcar EULA como ya aceptada
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OOBE" /v SetupDisplayedEula /t REG_DWORD /d 1 /f >nul 2>&1
echo         OK - EULA marcada como aceptada

:: Proteja su dispositivo - Omitir (valor 3 = no hacer nada)
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" /v ProtectYourPC /t REG_DWORD /d 3 /f >nul 2>&1
echo         OK - Proteccion de dispositivo: omitir

:: Saltar pantalla de vinculacion con telefono
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" /v SkipUserStatusPage /t REG_DWORD /d 1 /f >nul 2>&1
echo         OK - Pantalla de vinculacion: omitir

:: --- 5. Copiar unattend.xml a la ubicacion correcta ---
echo.
echo   [5/6] Copiando configuracion OOBE...

:: Crear directorio si no existe
if not exist "C:\Windows\Panther\Unattend" (
    mkdir "C:\Windows\Panther\Unattend" >nul 2>&1
)

:: Buscar autounattend.xml en varias ubicaciones
if exist "%SCRIPTPATH%autounattend.xml" (
    copy /y "%SCRIPTPATH%autounattend.xml" "C:\Windows\Panther\Unattend\unattend.xml" >nul 2>&1
    echo         OK - unattend.xml copiado desde %SCRIPTPATH%
) else if exist "%SCRIPTDRIVE%\autounattend.xml" (
    copy /y "%SCRIPTDRIVE%\autounattend.xml" "C:\Windows\Panther\Unattend\unattend.xml" >nul 2>&1
    echo         OK - unattend.xml copiado desde raiz USB
) else if exist "%SCRIPTDRIVE%\OOBE\autounattend.xml" (
    copy /y "%SCRIPTDRIVE%\OOBE\autounattend.xml" "C:\Windows\Panther\Unattend\unattend.xml" >nul 2>&1
    echo         OK - unattend.xml copiado desde carpeta OOBE
) else (
    echo         WARN - No se encontro autounattend.xml
    echo         Solo se aplican las configuraciones de registro
)

:: Verificar que se copio correctamente
if exist "C:\Windows\Panther\Unattend\unattend.xml" (
    echo         Verificado: unattend.xml existe en destino
) else (
    echo         WARN - unattend.xml NO se copio correctamente
)

:: --- 6. Reiniciar OOBE ---
echo.
echo   [6/6] Reiniciando OOBE...
echo.
echo   ============================================================
echo   La computadora se va a reiniciar en 5 segundos.
echo   Al volver, el OOBE saltara directo a crear usuario.
echo   ============================================================
echo.
timeout /t 5
shutdown /r /t 0
