# =============================================================================
# ACTUALIZADOR A WINDOWS 11 - COSMOS UPGRADE
# Hospital Civil de Guadalajara - Soporte Tecnico
# =============================================================================

# --- Auto-elevacion como Administrador ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# =============================================================================
# FUNCIONES DE VISUALIZACION
# =============================================================================

function Show-CosmosBanner {
    Clear-Host
    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *" -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "       ================================================" -ForegroundColor Cyan
    Write-Host "       |                                              |" -ForegroundColor Cyan
    Write-Host "       |   COSMOS UPGRADE - WINDOWS 11                |" -ForegroundColor Cyan
    Write-Host "       |                                              |" -ForegroundColor Cyan
    Write-Host "       |      Hospital Civil de Guadalajara           |" -ForegroundColor Cyan
    Write-Host "       |         Soporte Tecnico - iA                 |" -ForegroundColor Cyan
    Write-Host "       |                                              |" -ForegroundColor Cyan
    Write-Host "       ================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *" -ForegroundColor DarkYellow
    Write-Host ""
}

function Write-CosmosLog {
    param([string]$Mensaje, [string]$Tipo = "INFO")
    $Color = switch ($Tipo) {
        "OK"    { "Green" }
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "PHASE" { "Cyan" }
        default { "White" }
    }
    $Icon = switch ($Tipo) {
        "OK"    { "[OK]" }
        "ERROR" { "[X]" }
        "WARN"  { "[!]" }
        "PHASE" { "[*]" }
        default { "[-]" }
    }
    Write-Host "  $Icon $Mensaje" -ForegroundColor $Color
}

# =============================================================================
# VERIFICACIONES
# =============================================================================

function Get-WindowsVersion {
    try {
        $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $Build = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop
        return @{
            Caption = $OS.Caption
            Version = $Build.DisplayVersion
            Build = $Build.CurrentBuild
            IsWindows11 = [int]$Build.CurrentBuild -ge 22000
        }
    } catch {
        return @{ Caption = "Desconocido"; Version = ""; Build = "0"; IsWindows11 = $false }
    }
}

function Test-Windows11Compatibility {
    Write-Host ""
    Write-CosmosLog "FASE 1: Verificando compatibilidad con Windows 11..." "PHASE"
    Write-Host ""

    $Compatible = $true
    $Razones = @()

    # 1. Verificar si ya es Windows 11
    $WinVer = Get-WindowsVersion
    if ($WinVer.IsWindows11) {
        Write-CosmosLog "Este equipo YA tiene Windows 11 ($($WinVer.Version))" "OK"
        return @{ Compatible = $false; YaTieneWin11 = $true; Razones = @("Ya tiene Windows 11") }
    }
    Write-CosmosLog "Windows actual: $($WinVer.Caption) $($WinVer.Version)" "INFO"

    # 2. TPM 2.0
    try {
        $TPM = Get-CimInstance -Namespace "root\cimv2\Security\MicrosoftTpm" -ClassName Win32_Tpm -ErrorAction Stop
        if ($TPM.SpecVersion -like "2.0*") {
            Write-CosmosLog "TPM 2.0: Presente y compatible" "OK"
        } else {
            Write-CosmosLog "TPM: Version $($TPM.SpecVersion) (se requiere 2.0)" "WARN"
            $Razones += "TPM no es 2.0"
        }
    } catch {
        Write-CosmosLog "TPM: No detectado o no accesible" "WARN"
        $Razones += "TPM no detectado"
    }

    # 3. Secure Boot
    try {
        $SecureBoot = Confirm-SecureBootUEFI -ErrorAction Stop
        if ($SecureBoot) {
            Write-CosmosLog "Secure Boot: Habilitado" "OK"
        } else {
            Write-CosmosLog "Secure Boot: Deshabilitado" "WARN"
            $Razones += "Secure Boot deshabilitado"
        }
    } catch {
        Write-CosmosLog "Secure Boot: No se pudo verificar" "WARN"
    }

    # 4. RAM (minimo 4GB)
    $RAM = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB
    $RAMRounded = [math]::Round($RAM, 1)
    if ($RAM -ge 4) {
        Write-CosmosLog "RAM: $RAMRounded GB (minimo 4 GB)" "OK"
    } else {
        Write-CosmosLog "RAM: $RAMRounded GB (insuficiente)" "ERROR"
        $Compatible = $false
        $Razones += "RAM insuficiente"
    }

    # 5. Espacio en disco
    $Disco = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    $EspacioLibreGB = [math]::Round($Disco.FreeSpace / 1GB, 1)
    if ($EspacioLibreGB -ge 20) {
        Write-CosmosLog "Espacio libre: $EspacioLibreGB GB" "OK"
    } else {
        Write-CosmosLog "Espacio libre: $EspacioLibreGB GB (poco espacio)" "WARN"
        $Razones += "Poco espacio libre"
    }

    # 6. Procesador
    $CPU = Get-CimInstance Win32_Processor
    $Cores = $CPU.NumberOfCores
    $Speed = [math]::Round($CPU.MaxClockSpeed / 1000, 1)
    if ($Cores -ge 2 -and $Speed -ge 1) {
        Write-CosmosLog "CPU: $($CPU.Name.Trim()) - $Cores nucleos, $Speed GHz" "OK"
    } else {
        Write-CosmosLog "CPU: Puede no ser compatible" "WARN"
    }

    # 7. Arquitectura 64-bit
    if ([Environment]::Is64BitOperatingSystem) {
        Write-CosmosLog "Arquitectura: 64-bit" "OK"
    } else {
        Write-CosmosLog "Arquitectura: 32-bit (no compatible)" "ERROR"
        $Compatible = $false
        $Razones += "Sistema 32-bit"
    }

    return @{ Compatible = $Compatible; YaTieneWin11 = $false; Razones = $Razones }
}

# =============================================================================
# DESCARGA E INSTALACION
# =============================================================================

function Download-Windows11Assistant {
    Write-Host ""
    Write-CosmosLog "FASE 2: Descargando Windows 11 Installation Assistant..." "PHASE"
    Write-Host ""

    $DownloadUrl = "https://go.microsoft.com/fwlink/?linkid=2171764"
    $DestPath = "$env:TEMP\Windows11InstallationAssistant.exe"

    if (Test-Path $DestPath) {
        Remove-Item $DestPath -Force -ErrorAction SilentlyContinue
    }

    $MaxRetries = 3
    $Retry = 0
    $Success = $false

    while ($Retry -lt $MaxRetries -and -not $Success) {
        $Retry++
        Write-CosmosLog "Intento $Retry de $MaxRetries..." "INFO"

        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $DownloadUrl -OutFile $DestPath -UseBasicParsing -ErrorAction Stop

            if (Test-Path $DestPath) {
                $FileSize = (Get-Item $DestPath).Length / 1MB
                $FileSizeRounded = [math]::Round($FileSize, 1)
                if ($FileSize -gt 1) {
                    Write-CosmosLog "Descarga completada ($FileSizeRounded MB)" "OK"
                    $Success = $true
                } else {
                    throw "Archivo muy pequeno"
                }
            }
        } catch {
            Write-CosmosLog "Error en descarga: $($_.Exception.Message)" "WARN"
            if ($Retry -lt $MaxRetries) {
                Write-CosmosLog "Reintentando en 10 segundos..." "INFO"
                Start-Sleep -Seconds 10
            }
        }
    }

    if ($Success) {
        return $DestPath
    } else {
        Write-CosmosLog "No se pudo descargar despues de $MaxRetries intentos" "ERROR"
        return $null
    }
}

function Start-Windows11Upgrade {
    param([string]$AssistantPath)

    Write-Host ""
    Write-CosmosLog "FASE 3: Iniciando actualizacion a Windows 11..." "PHASE"
    Write-Host ""

    if (-not (Test-Path $AssistantPath)) {
        Write-CosmosLog "No se encontro el asistente de instalacion" "ERROR"
        return $false
    }

    Write-CosmosLog "Ejecutando Windows 11 Installation Assistant..." "INFO"
    Write-Host ""
    Write-Host "  =========================================================" -ForegroundColor Yellow
    Write-Host "  |  IMPORTANTE: La actualizacion puede tardar 30-60 min  |" -ForegroundColor Yellow
    Write-Host "  |  El equipo se reiniciara automaticamente              |" -ForegroundColor Yellow
    Write-Host "  |  NO apagues el equipo durante el proceso              |" -ForegroundColor Yellow
    Write-Host "  =========================================================" -ForegroundColor Yellow
    Write-Host ""

    try {
        $Arguments = "/quietinstall /skipeula /auto upgrade"
        Write-CosmosLog "Iniciando proceso de actualizacion..." "OK"
        $Process = Start-Process -FilePath $AssistantPath -ArgumentList $Arguments -PassThru -Wait

        if ($Process.ExitCode -eq 0) {
            Write-CosmosLog "El asistente se ejecuto correctamente" "OK"
            return $true
        } else {
            Write-CosmosLog "El asistente termino con codigo: $($Process.ExitCode)" "WARN"
            return $true
        }
    } catch {
        Write-CosmosLog "Error al ejecutar el asistente: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# =============================================================================
# EJECUCION PRINCIPAL
# =============================================================================

Show-CosmosBanner

# Verificar conexion a Internet
Write-CosmosLog "Verificando conexion a Internet..." "INFO"
$Internet = Test-Connection -ComputerName "8.8.8.8" -Count 2 -Quiet -ErrorAction SilentlyContinue
if (-not $Internet) {
    $Internet = Test-Connection -ComputerName "dns.google" -Count 2 -Quiet -ErrorAction SilentlyContinue
}
if (-not $Internet) {
    Write-CosmosLog "No hay conexion a Internet" "ERROR"
    Write-Host ""
    Write-Host "  Presiona cualquier tecla para salir..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}
Write-CosmosLog "Conexion a Internet: OK" "OK"

# FASE 1: Verificar compatibilidad
$Compat = Test-Windows11Compatibility

if ($Compat.YaTieneWin11) {
    Write-Host ""
    Write-Host "  =========================================================" -ForegroundColor Green
    Write-Host "  |                                                       |" -ForegroundColor Green
    Write-Host "  |   ESTE EQUIPO YA TIENE WINDOWS 11                     |" -ForegroundColor Green
    Write-Host "  |   No es necesario actualizar                          |" -ForegroundColor Green
    Write-Host "  |                                                       |" -ForegroundColor Green
    Write-Host "  =========================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Presiona cualquier tecla para salir..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

if (-not $Compat.Compatible) {
    Write-Host ""
    Write-CosmosLog "El equipo NO cumple los requisitos minimos" "ERROR"
    foreach ($Razon in $Compat.Razones) {
        Write-Host "    - $Razon" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "  Presiona cualquier tecla para salir..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Confirmacion
Write-Host ""
Write-Host "  =========================================================" -ForegroundColor Cyan
Write-Host "  |  El equipo es COMPATIBLE con Windows 11               |" -ForegroundColor Cyan
Write-Host "  |                                                       |" -ForegroundColor Cyan
Write-Host "  |  Deseas iniciar la actualizacion ahora?               |" -ForegroundColor Cyan
Write-Host "  |                                                       |" -ForegroundColor Cyan
Write-Host "  |  [S] Si, actualizar ahora                             |" -ForegroundColor Cyan
Write-Host "  |  [N] No, cancelar                                     |" -ForegroundColor Cyan
Write-Host "  |                                                       |" -ForegroundColor Cyan
Write-Host "  =========================================================" -ForegroundColor Cyan
Write-Host ""

$Respuesta = Read-Host "  Opcion (S/N)"
if ($Respuesta -notmatch "^[Ss]$") {
    Write-CosmosLog "Actualizacion cancelada por el usuario" "INFO"
    exit
}

# FASE 2: Descargar
$AssistantPath = Download-Windows11Assistant

if (-not $AssistantPath) {
    Write-Host ""
    Write-Host "  Presiona cualquier tecla para salir..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# FASE 3: Instalar
$Result = Start-Windows11Upgrade -AssistantPath $AssistantPath

if ($Result) {
    Write-Host ""
    Write-Host "  =========================================================" -ForegroundColor Green
    Write-Host "  |                                                       |" -ForegroundColor Green
    Write-Host "  |   ACTUALIZACION INICIADA                              |" -ForegroundColor Green
    Write-Host "  |                                                       |" -ForegroundColor Green
    Write-Host "  |   El equipo se reiniciara automaticamente             |" -ForegroundColor Green
    Write-Host "  |   cuando termine la actualizacion                     |" -ForegroundColor Green
    Write-Host "  |                                                       |" -ForegroundColor Green
    Write-Host "  |   NO APAGUES EL EQUIPO                                |" -ForegroundColor Green
    Write-Host "  |                                                       |" -ForegroundColor Green
    Write-Host "  =========================================================" -ForegroundColor Green
    Write-Host ""
} else {
    Write-CosmosLog "No se pudo iniciar la actualizacion" "ERROR"
}

Write-Host ""
Write-Host "  Presiona cualquier tecla para salir..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
