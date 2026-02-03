# =============================================================================
# CONFIGURADOR DE EQUIPOS - HOSPITAL CIVIL FAA
# =============================================================================
# Ejecutar como Administrador
# Version: 1.0
# Fecha: Enero 2026
# =============================================================================

#Requires -RunAsAdministrator

# Cargar configuracion
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$ScriptPath\Config.ps1"

# =============================================================================
# FUNCIONES DE UTILIDAD
# =============================================================================

function Write-Log {
    param(
        [string]$Mensaje,
        [string]$Tipo = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Tipo] $Mensaje"

    # Crear carpeta de logs si no existe
    if (-not (Test-Path $RutaLogs)) {
        New-Item -ItemType Directory -Path $RutaLogs -Force | Out-Null
    }

    $LogFile = Join-Path $RutaLogs "Config_$(Get-Date -Format 'yyyyMMdd').log"
    Add-Content -Path $LogFile -Value $LogEntry

    switch ($Tipo) {
        "ERROR" { Write-Host $LogEntry -ForegroundColor Red }
        "WARN"  { Write-Host $LogEntry -ForegroundColor Yellow }
        "OK"    { Write-Host $LogEntry -ForegroundColor Green }
        default { Write-Host $LogEntry -ForegroundColor Cyan }
    }
}

function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║                                                              ║" -ForegroundColor Cyan
    Write-Host "  ║     CONFIGURADOR DE EQUIPOS - HOSPITAL CIVIL FAA v1.0       ║" -ForegroundColor Cyan
    Write-Host "  ║     Coordinacion General de Informatica - Ext. 54425        ║" -ForegroundColor Cyan
    Write-Host "  ║                                                              ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Progress {
    param(
        [int]$Paso,
        [int]$Total,
        [string]$Descripcion
    )
    $Porcentaje = [math]::Round(($Paso / $Total) * 100)
    $Barra = "[" + ("█" * [math]::Floor($Porcentaje / 5)) + ("░" * (20 - [math]::Floor($Porcentaje / 5))) + "]"
    Write-Host "`n  $Barra $Porcentaje% - Paso $Paso de $Total" -ForegroundColor Yellow
    Write-Host "  >> $Descripcion" -ForegroundColor White
    Write-Host ""
}

# =============================================================================
# FUNCIONES DE EXTRACCION DE DATOS
# =============================================================================

function Get-DatosEquipo {
    Write-Log "Extrayendo datos del hardware..."

    $Datos = @{}

    try {
        # Numero de serie
        $Bios = Get-WmiObject Win32_BIOS
        $Datos.NumeroSerie = $Bios.SerialNumber

        # MAC Ethernet
        $AdaptadorEthernet = Get-WmiObject Win32_NetworkAdapter |
            Where-Object { $_.NetConnectionID -eq "Ethernet" -or $_.Name -like "*Ethernet*" -and $_.MACAddress } |
            Select-Object -First 1
        $Datos.MACEthernet = ($AdaptadorEthernet.MACAddress -replace ":", "")

        # MAC WiFi
        $AdaptadorWiFi = Get-WmiObject Win32_NetworkAdapter |
            Where-Object { $_.Name -like "*Wi-Fi*" -or $_.Name -like "*Wireless*" -and $_.MACAddress } |
            Select-Object -First 1
        $Datos.MACWiFi = ($AdaptadorWiFi.MACAddress -replace ":", "")

        # Product Key de Windows
        $DigitalProductId = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DigitalProductId
        $Datos.ProductKey = Get-WindowsProductKey

        # Fecha de fabricacion (aproximada desde BIOS)
        $Datos.FechaFabricacion = $Bios.ReleaseDate

        # UUID
        $CS = Get-WmiObject Win32_ComputerSystemProduct
        $Datos.UUID = $CS.UUID

        # Modelo
        $Sistema = Get-WmiObject Win32_ComputerSystem
        $Datos.Fabricante = $Sistema.Manufacturer
        $Datos.ModeloReal = $Sistema.Model

        Write-Log "Datos extraidos correctamente" "OK"
        Write-Log "  Serie: $($Datos.NumeroSerie)" "INFO"
        Write-Log "  MAC Ethernet: $($Datos.MACEthernet)" "INFO"
        Write-Log "  MAC WiFi: $($Datos.MACWiFi)" "INFO"

    } catch {
        Write-Log "Error extrayendo datos: $_" "ERROR"
    }

    return $Datos
}

function Get-WindowsProductKey {
    try {
        $Key = (Get-WmiObject -Query "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService" |
            Where-Object { $_.OA3xOriginalProductKey }).OA3xOriginalProductKey
        if ($Key) { return $Key }

        # Metodo alternativo
        $Output = & cscript //nologo "C:\Windows\System32\slmgr.vbs" /dli 2>$null
        $Match = [regex]::Match($Output, "Product Key: (.+)")
        if ($Match.Success) { return $Match.Groups[1].Value }

        return "No disponible"
    } catch {
        return "No disponible"
    }
}

# =============================================================================
# FUNCIONES DE CONFIGURACION DEL SISTEMA
# =============================================================================

function Set-NombreEquipo {
    param([string]$NumeroInventario)

    $NuevoNombre = "$PrefijoNombreEquipo$NumeroInventario"
    Write-Log "Renombrando equipo a: $NuevoNombre"

    try {
        Rename-Computer -NewName $NuevoNombre -Force -ErrorAction Stop
        Write-Log "Equipo renombrado correctamente" "OK"
        return $true
    } catch {
        Write-Log "Error al renombrar: $_" "ERROR"
        return $false
    }
}

function New-UsuarioSoporte {
    Write-Log "Creando usuario de soporte..."

    try {
        # Verificar si ya existe
        $Usuario = Get-LocalUser -Name $NombreUsuarioSoporte -ErrorAction SilentlyContinue

        if ($Usuario) {
            Write-Log "El usuario $NombreUsuarioSoporte ya existe" "WARN"
            return $true
        }

        # Crear usuario
        $SecurePassword = ConvertTo-SecureString $PasswordSoporte -AsPlainText -Force
        New-LocalUser -Name $NombreUsuarioSoporte `
                      -Password $SecurePassword `
                      -Description $DescripcionSoporte `
                      -PasswordNeverExpires `
                      -UserMayNotChangePassword `
                      -ErrorAction Stop

        # Agregar al grupo Administradores
        Add-LocalGroupMember -Group "Administradores" -Member $NombreUsuarioSoporte -ErrorAction SilentlyContinue

        Write-Log "Usuario $NombreUsuarioSoporte creado correctamente" "OK"
        return $true
    } catch {
        Write-Log "Error creando usuario: $_" "ERROR"
        return $false
    }
}

function Set-RedPrivada {
    Write-Log "Configurando red como privada..."

    try {
        # Obtener todos los perfiles de red activos
        $Perfiles = Get-NetConnectionProfile

        foreach ($Perfil in $Perfiles) {
            if ($Perfil.NetworkCategory -ne "Private") {
                Set-NetConnectionProfile -InterfaceIndex $Perfil.InterfaceIndex -NetworkCategory Private
                Write-Log "Red '$($Perfil.Name)' configurada como privada" "OK"
            }
        }

        # Habilitar descubrimiento de red
        netsh advfirewall firewall set rule group="Deteccion de redes" new enable=Yes 2>$null
        netsh advfirewall firewall set rule group="Network Discovery" new enable=Yes 2>$null

        # Habilitar compartir archivos e impresoras
        netsh advfirewall firewall set rule group="Compartir archivos e impresoras" new enable=Yes 2>$null
        netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes 2>$null

        Write-Log "Configuracion de red completada" "OK"
        return $true
    } catch {
        Write-Log "Error configurando red: $_" "ERROR"
        return $false
    }
}

function Set-HoraAutomatica {
    Write-Log "Configurando hora automatica..."

    try {
        # Configurar zona horaria
        Set-TimeZone -Id "Central Standard Time (Mexico)" -ErrorAction SilentlyContinue

        # Habilitar sincronizacion automatica
        Set-Service -Name w32time -StartupType Automatic
        Start-Service w32time -ErrorAction SilentlyContinue

        # Configurar servidor NTP
        w32tm /config /manualpeerlist:"time.windows.com" /syncfromflags:manual /reliable:yes /update
        w32tm /resync /force

        # Habilitar actualizacion automatica de hora en registro
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "Type" -Value "NTP"

        Write-Log "Hora automatica configurada" "OK"
        return $true
    } catch {
        Write-Log "Error configurando hora: $_" "ERROR"
        return $false
    }
}

function Set-TemaOscuro {
    Write-Log "Configurando tema oscuro..."

    try {
        $RegPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"

        # Tema oscuro para apps
        Set-ItemProperty -Path $RegPath -Name "AppsUseLightTheme" -Value 0 -Type DWord

        # Tema oscuro para sistema
        Set-ItemProperty -Path $RegPath -Name "SystemUsesLightTheme" -Value 0 -Type DWord

        # Color de enfasis automatico
        Set-ItemProperty -Path $RegPath -Name "ColorPrevalence" -Value 1 -Type DWord

        # Color de enfasis en barra de tareas
        $RegPathDWM = "HKCU:\SOFTWARE\Microsoft\Windows\DWM"
        Set-ItemProperty -Path $RegPathDWM -Name "ColorPrevalence" -Value 1 -Type DWord

        Write-Log "Tema oscuro configurado" "OK"
        return $true
    } catch {
        Write-Log "Error configurando tema: $_" "ERROR"
        return $false
    }
}

# =============================================================================
# FUNCIONES DE CONEXION A RED
# =============================================================================

function Connect-ServidoresRed {
    Write-Log "Conectando a servidores de red..."

    try {
        # Desconectar conexiones previas
        net use \\10.2.1.13 /delete /y 2>$null | Out-Null
        net use \\10.2.1.17 /delete /y 2>$null | Out-Null

        # Intentar usar credenciales guardadas en Credential Manager
        $Resultado1 = net use \\10.2.1.13\soportefaa 2>&1

        if ($LASTEXITCODE -ne 0) {
            # Si no hay credenciales guardadas, pedirlas
            Write-Log "Solicitando credenciales para servidor de software..." "INFO"
            Write-Host ""
            Write-Host "  Ingresa las credenciales de red:" -ForegroundColor Yellow
            $Usuario = Read-Host "  Usuario"
            $Password = Read-Host "  Password"

            $Resultado1 = net use \\10.2.1.13\soportefaa /user:$Usuario $Password /persistent:no 2>&1

            if ($LASTEXITCODE -eq 0) {
                # Guardar credenciales para futuro uso
                cmdkey /add:10.2.1.13 /user:$Usuario /pass:$Password | Out-Null
                Write-Log "Credenciales guardadas para 10.2.1.13" "OK"
            }
        }

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Conectado a \\10.2.1.13\soportefaa" "OK"
        } else {
            Write-Log "Error conectando a 10.2.1.13: $Resultado1" "ERROR"
            return $false
        }

        # Conectar a servidor de Dedalus (usa credenciales diferentes)
        $Resultado2 = net use \\10.2.1.17\distribucion /user:distribucion distribucion /persistent:no 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Conectado a \\10.2.1.17\distribucion" "OK"
        } else {
            Write-Log "Error conectando a 10.2.1.17: $Resultado2" "WARN"
            # No es critico si falla, continuamos
        }

        return $true
    } catch {
        Write-Log "Error conectando a servidores: $_" "ERROR"
        return $false
    }
}

function Get-StoredCredential {
    param([string]$Target)

    try {
        # Intentar usar cmdkey para obtener credenciales
        $Output = cmdkey /list:$Target 2>$null
        if ($Output -match $Target) {
            # Las credenciales existen, usar CredentialManager
            Add-Type -AssemblyName System.Web
            $Cred = [System.Net.CredentialCache]::DefaultNetworkCredentials
            return $Cred
        }
    } catch {}

    return $null
}

function Save-Credential {
    param(
        [string]$Target,
        [PSCredential]$Credential
    )

    try {
        $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))

        cmdkey /add:$Target /user:$($Credential.UserName) /pass:$Password
        Write-Log "Credenciales guardadas en Credential Manager" "OK"
    } catch {
        Write-Log "No se pudieron guardar las credenciales: $_" "WARN"
    }
}

# =============================================================================
# FUNCIONES DE INSTALACION
# =============================================================================

function Uninstall-Office365 {
    Write-Log "Buscando Office 365 para desinstalar..."

    try {
        # Buscar instalaciones de Office 365/Microsoft 365
        $Office365 = Get-WmiObject -Class Win32_Product |
            Where-Object { $_.Name -like "*Microsoft 365*" -or $_.Name -like "*Office 365*" }

        if ($Office365) {
            foreach ($App in $Office365) {
                Write-Log "Desinstalando: $($App.Name)"
                $App.Uninstall() | Out-Null
            }
            Write-Log "Office 365 desinstalado" "OK"
        } else {
            # Intentar con la herramienta de desinstalacion de Office
            $SaRAPath = "$env:TEMP\SaRA_Office.exe"
            # Si tienes la herramienta SaRA, se puede usar aqui
            Write-Log "No se encontro Office 365 instalado" "INFO"
        }

        return $true
    } catch {
        Write-Log "Error desinstalando Office 365: $_" "ERROR"
        return $false
    }
}

function Install-Office2007 {
    Write-Log "Instalando Office 2007 (Word, Excel, PowerPoint)..."

    try {
        # Buscar el setup.exe de Office
        $SetupPath = Get-ChildItem -Path $RutaOffice2007 -Filter "setup.exe" -Recurse |
            Select-Object -First 1 -ExpandProperty FullName

        if (-not $SetupPath) {
            $SetupPath = Get-ChildItem -Path $RutaOffice2007 -Filter "*.exe" -Recurse |
                Select-Object -First 1 -ExpandProperty FullName
        }

        if ($SetupPath) {
            # Crear archivo de configuracion para instalacion silenciosa
            $ConfigXML = @"
<Configuration Product="Enterprise">
    <Display Level="none" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" />
    <PIDKEY Value="$($SerialOffice2007 -replace '-','')" />
    <OptionState Id="WORDFiles" State="local" Children="force" />
    <OptionState Id="EXCELFiles" State="local" Children="force" />
    <OptionState Id="PPTFiles" State="local" Children="force" />
    <OptionState Id="OUTLOOKFiles" State="absent" Children="force" />
    <OptionState Id="ACCESSFiles" State="absent" Children="force" />
    <OptionState Id="PUBFiles" State="absent" Children="force" />
    <OptionState Id="ONOTEFiles" State="absent" Children="force" />
    <OptionState Id="InfoijPath" State="absent" Children="force" />
    <OptionState Id="GROOVEFiles" State="absent" Children="force" />
</Configuration>
"@
            $ConfigPath = "$env:TEMP\Office2007Config.xml"
            $ConfigXML | Out-File -FilePath $ConfigPath -Encoding UTF8

            # Ejecutar instalacion silenciosa
            Start-Process -FilePath $SetupPath -ArgumentList "/config `"$ConfigPath`"" -Wait -NoNewWindow

            Write-Log "Office 2007 instalado correctamente" "OK"
        } else {
            Write-Log "No se encontro el instalador de Office 2007" "ERROR"
            return $false
        }

        return $true
    } catch {
        Write-Log "Error instalando Office 2007: $_" "ERROR"
        return $false
    }
}

function Install-ESETAntivirus {
    Write-Log "Instalando ESET Antivirus..."

    try {
        # Ruta directa al instalador
        if (Test-Path $RutaAntivirus) {
            Write-Log "Ejecutando: $RutaAntivirus"
            # ESET PROTECT Installer - instalacion silenciosa
            Start-Process -FilePath $RutaAntivirus -ArgumentList "--silent --accepteula" -Wait -NoNewWindow
            Write-Log "ESET Antivirus instalado correctamente" "OK"
        } else {
            Write-Log "No se encontro instalador de ESET en: $RutaAntivirus" "ERROR"
            return $false
        }

        return $true
    } catch {
        Write-Log "Error instalando ESET: $_" "ERROR"
        return $false
    }
}

function Install-Chrome {
    Write-Log "Instalando Google Chrome..."

    try {
        # Verificar si ya esta instalado
        $Chrome = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe" -ErrorAction SilentlyContinue

        if ($Chrome) {
            Write-Log "Google Chrome ya esta instalado" "INFO"
            return $true
        }

        # Usar instalador del servidor
        if (Test-Path $RutaChrome) {
            Write-Log "Ejecutando instalador desde servidor..."
            Start-Process -FilePath $RutaChrome -ArgumentList "/silent /install" -Wait -NoNewWindow
            Write-Log "Google Chrome instalado correctamente" "OK"
        } else {
            # Descargar si no esta en el servidor
            Write-Log "Descargando Chrome desde internet..."
            $ChromeUrl = "https://dl.google.com/chrome/install/GoogleChromeStandaloneEnterprise64.msi"
            $ChromePath = "$env:TEMP\ChromeSetup.msi"
            Invoke-WebRequest -Uri $ChromeUrl -OutFile $ChromePath -UseBasicParsing
            Start-Process msiexec -ArgumentList "/i `"$ChromePath`" /quiet /norestart" -Wait
            Write-Log "Google Chrome instalado correctamente" "OK"
        }

        return $true
    } catch {
        Write-Log "Error instalando Chrome: $_" "ERROR"
        return $false
    }
}

function Install-AcrobatReader {
    Write-Log "Instalando Acrobat Reader DC..."

    try {
        # Ruta directa al instalador
        if (Test-Path $RutaAcrobat) {
            Write-Log "Ejecutando: $RutaAcrobat"
            # Instalacion silenciosa de Acrobat Reader
            Start-Process -FilePath $RutaAcrobat -ArgumentList "/sAll /rs /msi EULA_ACCEPT=YES" -Wait -NoNewWindow
            Write-Log "Acrobat Reader instalado correctamente" "OK"
        } else {
            Write-Log "No se encontro instalador de Acrobat Reader en: $RutaAcrobat" "ERROR"
            return $false
        }

        return $true
    } catch {
        Write-Log "Error instalando Acrobat Reader: $_" "ERROR"
        return $false
    }
}

function Install-DotNetFramework35 {
    Write-Log "Instalando .NET Framework 3.5..."

    try {
        # Verificar si ya esta instalado
        $DotNet35 = Get-WindowsOptionalFeature -Online -FeatureName "NetFx3" -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq "Enabled" }

        if ($DotNet35) {
            Write-Log ".NET Framework 3.5 ya esta instalado" "INFO"
            return $true
        }

        # Primero intentar con el instalador del servidor (mas rapido)
        if (Test-Path $RutaDotNet35) {
            Write-Log "Instalando desde servidor: $RutaDotNet35"
            Start-Process -FilePath $RutaDotNet35 -ArgumentList "/q /norestart" -Wait -NoNewWindow
            Write-Log ".NET Framework 3.5 instalado desde servidor" "OK"
            return $true
        }

        # Si no esta en servidor, usar DISM (requiere internet)
        Write-Log "Habilitando .NET Framework 3.5 via DISM (puede tardar)..."
        Enable-WindowsOptionalFeature -Online -FeatureName "NetFx3" -All -NoRestart

        Write-Log ".NET Framework 3.5 instalado correctamente" "OK"
        return $true
    } catch {
        Write-Log "Error instalando .NET Framework 3.5: $_" "ERROR"
        return $false
    }
}

function Install-Dedalus {
    Write-Log "Instalando Dedalus (Expediente Electronico)..."

    try {
        # Copiar sincronizador
        $DestinoSync = "C:\Dedalus"
        if (-not (Test-Path $DestinoSync)) {
            New-Item -ItemType Directory -Path $DestinoSync -Force | Out-Null
        }

        # Copiar archivos del sincronizador
        Copy-Item -Path "$RutaSincronizador\*" -Destination $DestinoSync -Recurse -Force

        # Ejecutar netlogon6.bat si existe
        $Netlogon = "$RutaSincronizador\netlogon6.bat"
        if (Test-Path $Netlogon) {
            Start-Process cmd -ArgumentList "/c `"$Netlogon`"" -Wait -NoNewWindow
        }

        # Copiar sync_xhis6.bat al destino
        $SyncBat = "$RutaSincronizador\sync_xhis6.bat"
        if (Test-Path $SyncBat) {
            Copy-Item -Path $SyncBat -Destination "C:\Dedalus\sync_xhis6.bat" -Force
        }

        Write-Log "Dedalus configurado correctamente" "OK"
        return $true
    } catch {
        Write-Log "Error instalando Dedalus: $_" "ERROR"
        return $false
    }
}

function Copy-AccesosDirectos {
    Write-Log "Copiando accesos directos al escritorio..."

    try {
        $DesktopPublic = [Environment]::GetFolderPath("CommonDesktopDirectory")

        # Copiar accesos desde el servidor
        if (Test-Path $RutaAccesos) {
            Get-ChildItem -Path $RutaAccesos -Filter "*.lnk" | ForEach-Object {
                Copy-Item -Path $_.FullName -Destination $DesktopPublic -Force
                Write-Log "Copiado: $($_.Name)" "INFO"
            }
            Write-Log "Accesos directos copiados" "OK"
        } else {
            Write-Log "No se encontro la carpeta de accesos directos" "WARN"
        }

        return $true
    } catch {
        Write-Log "Error copiando accesos: $_" "ERROR"
        return $false
    }
}

# =============================================================================
# FUNCION DE GENERACION DE REPORTE
# =============================================================================

function New-ReporteEquipo {
    param(
        [string]$NumeroInventario,
        [hashtable]$DatosHW
    )

    Write-Log "Generando reporte del equipo..."

    $Reporte = @{
        FechaRegistro = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        NumeroInventarioST = $NumeroInventario
        NombreEquipo = "$PrefijoNombreEquipo$NumeroInventario"
        NumeroSerie = $DatosHW.NumeroSerie
        MACEthernet = $DatosHW.MACEthernet
        MACWiFi = $DatosHW.MACWiFi
        ProductKey = $DatosHW.ProductKey
        UUID = $DatosHW.UUID
        Fabricante = $DatosHW.Fabricante
        Modelo = $DatosHW.ModeloReal
        Especificaciones = $EspecsEquipo
    }

    # Guardar como JSON
    if (-not (Test-Path $RutaReportesLocal)) {
        New-Item -ItemType Directory -Path $RutaReportesLocal -Force | Out-Null
    }

    $ArchivoReporte = Join-Path $RutaReportesLocal "Equipo_$NumeroInventario`_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $Reporte | ConvertTo-Json -Depth 5 | Out-File -FilePath $ArchivoReporte -Encoding UTF8

    Write-Log "Reporte guardado en: $ArchivoReporte" "OK"

    return $Reporte
}

# =============================================================================
# FUNCION PRINCIPAL
# =============================================================================

function Start-ConfiguracionEquipo {
    Show-Banner

    $TotalPasos = 14
    $PasoActual = 0

    # Pedir numero de inventario
    Write-Host ""
    Write-Host "  Ingresa el numero de inventario de Soporte Tecnico" -ForegroundColor Yellow
    Write-Host "  (Es el numero de 5 digitos de la etiqueta azul)" -ForegroundColor Gray
    Write-Host ""
    $NumeroInventario = Read-Host "  Numero de inventario ST"

    if ($NumeroInventario -notmatch '^\d{5}$') {
        Write-Log "El numero de inventario debe ser de 5 digitos" "ERROR"
        return
    }

    Write-Log "Iniciando configuracion para equipo: $NumeroInventario"
    Write-Host ""

    # Paso 1: Extraer datos del hardware
    $PasoActual++
    Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Extrayendo datos del hardware"
    $DatosHW = Get-DatosEquipo

    # Paso 2: Renombrar equipo
    $PasoActual++
    Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Renombrando equipo"
    Set-NombreEquipo -NumeroInventario $NumeroInventario

    # Paso 3: Crear usuario de soporte
    $PasoActual++
    Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Creando usuario de soporte"
    New-UsuarioSoporte

    # Paso 4: Configurar red privada
    $PasoActual++
    Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Configurando red como privada"
    Set-RedPrivada

    # Paso 5: Configurar hora automatica
    $PasoActual++
    Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Configurando hora automatica"
    Set-HoraAutomatica

    # Paso 6: Configurar tema oscuro
    $PasoActual++
    Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Configurando tema oscuro"
    Set-TemaOscuro

    # Paso 7: Conectar a servidores de red
    $PasoActual++
    Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Conectando a servidores de red"
    $Conectado = Connect-ServidoresRed

    if ($Conectado) {
        # Paso 8: Desinstalar Office 365
        $PasoActual++
        Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Desinstalando Office 365"
        Uninstall-Office365

        # Paso 9: Instalar .NET Framework 3.5
        $PasoActual++
        Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Instalando .NET Framework 3.5"
        Install-DotNetFramework35

        # Paso 10: Instalar Office 2007
        $PasoActual++
        Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Instalando Office 2007"
        Install-Office2007

        # Paso 11: Instalar ESET Antivirus
        $PasoActual++
        Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Instalando ESET Antivirus"
        Install-ESETAntivirus

        # Paso 12: Instalar Chrome
        $PasoActual++
        Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Instalando Google Chrome"
        Install-Chrome

        # Paso 13: Instalar Acrobat Reader
        $PasoActual++
        Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Instalando Acrobat Reader"
        Install-AcrobatReader

        # Paso 14: Copiar accesos directos
        $PasoActual++
        Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Copiando accesos directos"
        Copy-AccesosDirectos

        # Instalar Dedalus
        # Show-Progress -Paso $PasoActual -Total $TotalPasos -Descripcion "Instalando Dedalus"
        # Install-Dedalus
    }

    # Generar reporte
    Write-Host ""
    Write-Host "  ══════════════════════════════════════════════════════════" -ForegroundColor Green
    $Reporte = New-ReporteEquipo -NumeroInventario $NumeroInventario -DatosHW $DatosHW

    # Mostrar resumen
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "  ║              CONFIGURACION COMPLETADA                        ║" -ForegroundColor Green
    Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Resumen del equipo:" -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────" -ForegroundColor Gray
    Write-Host "  Nombre:        $PrefijoNombreEquipo$NumeroInventario" -ForegroundColor White
    Write-Host "  No. Serie:     $($DatosHW.NumeroSerie)" -ForegroundColor White
    Write-Host "  MAC Ethernet:  $($DatosHW.MACEthernet)" -ForegroundColor White
    Write-Host "  MAC WiFi:      $($DatosHW.MACWiFi)" -ForegroundColor White
    Write-Host "  Product Key:   $($DatosHW.ProductKey)" -ForegroundColor White
    Write-Host ""
    Write-Host "  IMPORTANTE: Reinicia el equipo para aplicar todos los cambios" -ForegroundColor Yellow
    Write-Host ""

    # Preguntar si desea reiniciar
    $Reiniciar = Read-Host "  Deseas reiniciar ahora? (S/N)"
    if ($Reiniciar -eq "S" -or $Reiniciar -eq "s") {
        Write-Log "Reiniciando equipo..."
        Restart-Computer -Force
    }
}

# =============================================================================
# PUNTO DE ENTRADA
# =============================================================================

# Ejecutar si se llama directamente
if ($MyInvocation.InvocationName -ne '.') {
    Start-ConfiguracionEquipo
}
