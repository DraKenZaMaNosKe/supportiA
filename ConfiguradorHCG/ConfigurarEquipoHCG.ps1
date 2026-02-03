# =============================================================================
# CONFIGURADOR DE EQUIPOS - HOSPITAL CIVIL FAA v4.0
# =============================================================================
# Se eleva automaticamente como Administrador si no lo es
# =============================================================================

# --- Auto-elevacion como Administrador ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# --- CONFIGURACION ---
$Servidor = "10.2.1.13"
$Usuario = "2010201"
$Password = "7v3l73v37nG06"

# Ruta base del pack de instalacion
$RutaBase = "\\$Servidor\soportefaa\pack_installer_iA"
$RutaAccesos = "$RutaBase\accesos_directos"
$RutaAcrobat = "$RutaBase\acrobat_reader"
$RutaAntivirus = "$RutaBase\antivirus"
$RutaChrome = "$RutaBase\chrome"
$RutaDedalus = "$RutaBase\dedalus_expedienteclinico"
$RutaOffice = "$RutaBase\office"
$RutaWallpaper = "$RutaBase\wallpaper"
$RutaDotNet = "$RutaBase\netframework3_5\sxs"
$RutaWinRAR = "$RutaBase\winrar_licence"

$ZonaHoraria = "Central Standard Time (Mexico)"
$UsuarioSoporte = "Soporte"
$PasswordSoporte = "*TIsoporte"
$RutaLogs = "C:\HCG_Logs"

$GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"

# Variable global para tracking de software instalado
$Script:SoftwareInstalado = @()
$Script:UsuarioOriginal = $env:USERNAME

# --- FUNCIONES AUXILIARES ---

function Write-Log {
    param([string]$Mensaje, [string]$Tipo = "INFO")
    $Icon = switch ($Tipo) {
        "OK"    { [char]0x2605 }  # Estrella dorada
        "ERROR" { [char]0x2716 }  # Cruz roja
        "WARN"  { [char]0x26A0 }  # Triangulo alerta
        default { [char]0x2192 }  # Flecha cosmica
    }
    $Color = switch ($Tipo) { "OK" { "Green" } "ERROR" { "Red" } "WARN" { "Yellow" } default { "Cyan" } }
    Write-Host "  $Icon [$Tipo] $Mensaje" -ForegroundColor $Color
    if (-not (Test-Path $RutaLogs)) { New-Item -ItemType Directory -Path $RutaLogs -Force | Out-Null }
    Add-Content -Path "$RutaLogs\config.log" -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Tipo] $Mensaje"

    # Sonidos sutiles segun tipo
    try {
        switch ($Tipo) {
            "OK"    { [Console]::Beep(880, 60); [Console]::Beep(1100, 60) }
            "ERROR" { [Console]::Beep(330, 150); [Console]::Beep(220, 200) }
            "WARN"  { [Console]::Beep(440, 100) }
        }
    } catch {}
}

function Show-Banner {
    Clear-Host
    $OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

    # Animacion de encendido del cosmo
    $CosmosChars = @(".", "*", "+", "x", "*", "+", ".", "*")
    for ($i = 0; $i -lt 3; $i++) {
        $Line = "  "
        for ($j = 0; $j -lt 60; $j++) {
            $Line += $CosmosChars[(Get-Random -Maximum $CosmosChars.Count)]
        }
        Write-Host $Line -ForegroundColor DarkYellow -NoNewline
        Start-Sleep -Milliseconds 80
        Write-Host "`r" -NoNewline
    }

    Clear-Host
    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  ." -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "                    /\        /\" -ForegroundColor DarkYellow
    Write-Host "                   /  \  /\  /  \" -ForegroundColor DarkYellow
    Write-Host "                  / /\ \/ /\/ /\ \" -ForegroundColor DarkYellow
    Write-Host "                 / /  \  / /\/  \ \" -ForegroundColor DarkYellow
    Write-Host "                /_/    \/\/ /    \_\" -ForegroundColor DarkYellow
    Write-Host "                \  PEGASUS  \/    /" -ForegroundColor Cyan
    Write-Host "                 \   / /\  /\   /" -ForegroundColor DarkYellow
    Write-Host "                  \_/ /  \/  \_/" -ForegroundColor DarkYellow
    Write-Host "                     /   /\" -ForegroundColor DarkYellow
    Write-Host "                    /   /  \" -ForegroundColor DarkYellow
    Write-Host "                   /   /    \" -ForegroundColor DarkYellow
    Write-Host "                  /___/______\" -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  ." -ForegroundColor DarkYellow
    Write-Host ""

    # Titulo epico con animacion progresiva
    $Title = "     CONFIGURADOR COSMICO DE EQUIPOS v4.0"
    foreach ($char in $Title.ToCharArray()) {
        Write-Host $char -NoNewline -ForegroundColor Cyan
        Start-Sleep -Milliseconds 15
    }
    Write-Host ""

    $Subtitle = "     Hospital Civil FAA - Caballeros de Informatica"
    foreach ($char in $Subtitle.ToCharArray()) {
        Write-Host $char -NoNewline -ForegroundColor Magenta
        Start-Sleep -Milliseconds 10
    }
    Write-Host ""

    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  ." -ForegroundColor DarkYellow
    Write-Host "           Ext. 54425 - Enciende tu cosmo!  " -ForegroundColor DarkYellow
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  ." -ForegroundColor DarkYellow
    Write-Host ""

    # Sonido de inicio cosmico
    try {
        [Console]::Beep(523, 100)  # Do
        [Console]::Beep(659, 100)  # Mi
        [Console]::Beep(784, 100)  # Sol
        [Console]::Beep(1047, 200) # Do alto
    } catch {}
}

function Write-StepHeader {
    param([int]$Step, [int]$Total = 25, [string]$Title)
    $Stars = [char]0x2605  # Estrella
    $Cosmos = [char]0x2734 # Estrella de 8 puntas
    $Filled = [char]0x2593 # Bloque lleno
    $Empty  = [char]0x2591 # Bloque vacio

    Write-Host ""

    # Barra de progreso cosmica
    $Pct = [math]::Floor(($Step / $Total) * 30)
    $Bar = ""
    for ($i = 0; $i -lt 30; $i++) {
        if ($i -lt $Pct) { $Bar += $Filled } else { $Bar += $Empty }
    }
    Write-Host "  $Cosmos " -NoNewline -ForegroundColor Magenta
    Write-Host "[$Bar]" -NoNewline -ForegroundColor DarkYellow
    Write-Host " $([math]::Floor(($Step / $Total) * 100))%" -ForegroundColor DarkYellow

    # Titulo del paso
    Write-Host "  $Stars " -NoNewline -ForegroundColor DarkYellow
    Write-Host "[$Step/$Total] " -NoNewline -ForegroundColor DarkYellow
    Write-Host "$Title" -ForegroundColor Cyan

    # Separador cosmico
    Write-Separator
}

function Write-Separator {
    $Star = [char]0x2606  # Estrella vacia
    $Dot  = [char]0x00B7  # Punto medio
    $Sep = "  "
    for ($i = 0; $i -lt 15; $i++) {
        if ($i % 3 -eq 0) { $Sep += "$Star " } else { $Sep += "$Dot$Dot " }
    }
    Write-Host $Sep -ForegroundColor DarkGray
}

function Play-StepSound {
    try {
        [Console]::Beep(659, 80)   # Mi
        [Console]::Beep(784, 80)   # Sol
        [Console]::Beep(1047, 120) # Do alto
    } catch {}
}

function Play-VictorySound {
    try {
        # Fanfarria de victoria tipo Saint Seiya
        [Console]::Beep(523, 120)  # Do
        [Console]::Beep(659, 120)  # Mi
        [Console]::Beep(784, 120)  # Sol
        [Console]::Beep(1047, 200) # Do alto
        Start-Sleep -Milliseconds 50
        [Console]::Beep(988, 120)  # Si
        [Console]::Beep(1047, 120) # Do alto
        [Console]::Beep(1175, 120) # Re alto
        [Console]::Beep(1319, 300) # Mi alto
        Start-Sleep -Milliseconds 100
        [Console]::Beep(1568, 150) # Sol alto
        [Console]::Beep(1319, 150) # Mi alto
        [Console]::Beep(1568, 400) # Sol alto (final)
    } catch {}
}

function Show-ProgressCosmos {
    param([int]$Step, [int]$Total = 25)
    $Pct = [math]::Floor(($Step / $Total) * 100)
    $title = "Cosmo: $Pct% - Paso $Step/$Total"
    $Host.UI.RawUI.WindowTitle = "$([char]0x2605) $title $([char]0x2605)"
}

function Show-CosmosAnimation {
    param([string]$Message = "Encendiendo el cosmo...")
    $Frames = @(
        "    .  *  .     ",
        "   * . + . *    ",
        "  . + * + .  *  ",
        " * . + * + . *  ",
        "  . + * + .  *  ",
        "   * . + . *    ",
        "    .  *  .     "
    )
    foreach ($frame in $Frames) {
        Write-Host "`r  $frame $Message" -NoNewline -ForegroundColor DarkYellow
        Start-Sleep -Milliseconds 80
    }
    Write-Host ""
}

function Find-Installer {
    param([string]$Carpeta, [string]$Filtro = "*.exe")
    if (Test-Path $Carpeta) {
        $Archivo = Get-ChildItem -Path $Carpeta -Filter $Filtro -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($Archivo) { return $Archivo.FullName }
    }
    return $null
}

function Pin-ToTaskbar {
    param([string]$ShortcutPath, [string]$AppName = "App")

    try {
        # Metodo para Windows 10/11: Copiar a la carpeta de TaskBar
        $TaskBarFolder = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

        if (-not (Test-Path $TaskBarFolder)) {
            New-Item -ItemType Directory -Path $TaskBarFolder -Force | Out-Null
        }

        if (Test-Path $ShortcutPath) {
            $FileName = Split-Path $ShortcutPath -Leaf
            $Destino = "$TaskBarFolder\$FileName"
            Copy-Item -Path $ShortcutPath -Destination $Destino -Force
            Write-Log "$AppName anclado a barra de tareas" "OK"
        }

        # Metodo alternativo usando Shell
        $Shell = New-Object -ComObject Shell.Application
        $Folder = $Shell.Namespace((Split-Path $ShortcutPath -Parent))
        $Item = $Folder.ParseName((Split-Path $ShortcutPath -Leaf))

        if ($Item) {
            $Verbs = $Item.Verbs()
            foreach ($Verb in $Verbs) {
                if ($Verb.Name -match "Anclar a la barra de tareas|Pin to taskbar|Anclar al inicio") {
                    $Verb.DoIt()
                    break
                }
            }
        }
    } catch {
        Write-Log "No se pudo anclar $AppName a la barra de tareas" "WARN"
    }
}

function Create-TaskbarShortcut {
    param(
        [string]$TargetPath,
        [string]$ShortcutName,
        [string]$IconPath = "",
        [string]$Arguments = ""
    )

    $TaskBarFolder = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    $ShortcutPath = "$TaskBarFolder\$ShortcutName.lnk"

    if (-not (Test-Path $TaskBarFolder)) {
        New-Item -ItemType Directory -Path $TaskBarFolder -Force | Out-Null
    }

    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    if ($Arguments) { $Shortcut.Arguments = $Arguments }
    if ($IconPath -and (Test-Path $IconPath)) { $Shortcut.IconLocation = $IconPath }
    $Shortcut.Save()

    return $ShortcutPath
}

function Get-DatosEquipo {
    $Bios = Get-WmiObject Win32_BIOS
    $MACEth = (Get-WmiObject Win32_NetworkAdapter | Where-Object { $_.NetConnectionID -eq "Ethernet" -and $_.MACAddress } | Select-Object -First 1).MACAddress -replace ":", ""
    $MACWifi = (Get-WmiObject Win32_NetworkAdapter | Where-Object { $_.NetConnectionID -eq "Wi-Fi" -and $_.MACAddress } | Select-Object -First 1).MACAddress -replace ":", ""
    $ProductKey = ""
    try { $ProductKey = (Get-WmiObject -Query "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService" | Where-Object { $_.OA3xOriginalProductKey }).OA3xOriginalProductKey } catch {}
    return @{ Serie = $Bios.SerialNumber; MACEthernet = $MACEth; MACWiFi = $MACWifi; ProductKey = $ProductKey }
}

# --- FUNCIONES DE GOOGLE SHEETS ---

function Send-DatosInicio {
    param([string]$InvST, [hashtable]$Datos)

    Write-StepHeader -Step 1 -Title "REGISTRANDO INICIO EN GOOGLE SHEETS"
    Show-ProgressCosmos -Step 1
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $FechaHoy = Get-Date -Format "dd/MM/yyyy"
    $FechaGarantia = (Get-Date).AddYears(3).ToString("dd/MM/yyyy")

    $CPU = (Get-WmiObject Win32_Processor | Select-Object -First 1).Name -replace '\(R\)|\(TM\)|CPU|@.*', '' -replace '\s+', ' '
    $Nucleos = (Get-WmiObject Win32_Processor | Select-Object -First 1).NumberOfCores
    $RAMTotal = [math]::Round((Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 0)
    $Disco = Get-WmiObject Win32_DiskDrive | Select-Object -First 1
    $DiscoGB = [math]::Round($Disco.Size / 1GB, 0)
    $DiscoTipo = if ($Disco.Model -like "*SSD*" -or $Disco.Model -like "*NVMe*") { "SSD" } else { "HDD" }

    $Body = @{
        Accion = "crear"
        Fecha = $FechaHoy
        InvST = $InvST
        Serie = $Datos.Serie
        Marca = "Lenovo"
        Modelo = "ThinkCentre M70s Gen 5"
        Procesador = $CPU.Trim()
        Nucleos = $Nucleos
        RAM = $RAMTotal
        Disco = $DiscoGB
        DiscoTipo = $DiscoTipo
        MACEthernet = $Datos.MACEthernet
        MACWiFi = $Datos.MACWiFi
        ProductKey = $Datos.ProductKey
        FechaFab = $FechaHoy
        Garantia = $FechaGarantia
    } | ConvertTo-Json

    try {
        Write-Host "  Enviando datos del equipo..." -ForegroundColor Cyan
        $Response = Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 60

        if ($Response.status -eq "OK") {
            Write-Host ""
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |       EQUIPO REGISTRADO EN SHEETS        |" -ForegroundColor Green
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |  Inventario:  $InvST" -ForegroundColor White
            Write-Host "  |  No. Serie:   $($Datos.Serie)" -ForegroundColor White
            Write-Host "  |  FILA:        $($Response.row)" -ForegroundColor Yellow
            if ($Response.faa -and $Response.faa -ne "") {
                Write-Host "  |  FAA:         $($Response.faa)" -ForegroundColor Green
            }
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host ""
            $Script:FilaRegistro = $Response.row
            return $true
        }
    } catch {
        Write-Host "  [ERROR] No se pudo registrar: $($_.Exception.Message)" -ForegroundColor Red
    }
    return $false
}

function Send-DatosFin {
    param([string]$InvST)

    Write-StepHeader -Step 24 -Title "ACTUALIZANDO GOOGLE SHEETS - COMPLETADO"
    Show-ProgressCosmos -Step 24
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $SoftwareList = if ($Script:SoftwareInstalado.Count -gt 0) { $Script:SoftwareInstalado -join ", " } else { "Configuracion completa" }

    $Body = @{
        Accion = "actualizar"
        InvST = $InvST
        SoftwareInstalado = $SoftwareList
    } | ConvertTo-Json

    try {
        $Response = Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 60
        if ($Response.status -eq "OK") {
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |     EQUIPO MARCADO COMO ACTIVO           |" -ForegroundColor Green
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            return $true
        }
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
    return $false
}

function Get-SoftwareInfo {
    $OS = Get-WmiObject Win32_OperatingSystem
    $WindowsVersion = $OS.Caption -replace "Microsoft ", ""
    $WindowsBuild = $OS.BuildNumber
    $LicenseStatus = (Get-WmiObject -Query "SELECT LicenseStatus FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL AND Name LIKE 'Windows%'" | Select-Object -First 1).LicenseStatus
    $WindowsActivado = if ($LicenseStatus -eq 1) { "Si" } else { "No" }
    $ProductKey = ""
    try { $ProductKey = (Get-WmiObject -Query "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService" | Where-Object { $_.OA3xOriginalProductKey }).OA3xOriginalProductKey } catch {}

    $OfficeVersion = "No"
    if (Test-Path "C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.EXE") { $OfficeVersion = "2007" }
    elseif (Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE") { $OfficeVersion = "365/2016+" }

    $ChromeVersion = "No"
    if (Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") {
        try { $ChromeVersion = (Get-Item "C:\Program Files\Google\Chrome\Application\chrome.exe").VersionInfo.FileVersion } catch { $ChromeVersion = "Si" }
    }

    $AcrobatVersion = "No"
    $AcrobatPaths = @("C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe", "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe")
    foreach ($path in $AcrobatPaths) { if (Test-Path $path) { $AcrobatVersion = "Si"; break } }

    $DotNet35 = if ((Get-WindowsOptionalFeature -Online -FeatureName "NetFx3" -ErrorAction SilentlyContinue).State -eq "Enabled") { "Si" } else { "No" }
    $DedalusInstalado = if (Test-Path "C:\Dedalus") { "Si" } else { "No" }
    $ESETInstalado = if (Get-Service -Name "ekrn" -ErrorAction SilentlyContinue) { "Si" } else { "No" }
    $WinRARInstalado = if (Test-Path "C:\Program Files\WinRAR\WinRAR.exe") { "Si" } else { "No" }

    return @{
        WindowsVersion = $WindowsVersion; WindowsBuild = $WindowsBuild; WindowsActivado = $WindowsActivado
        ProductKey = $ProductKey; NombreEquipo = $env:COMPUTERNAME; UsuarioWindows = $env:USERNAME
        Office = $OfficeVersion; Chrome = $ChromeVersion; Acrobat = $AcrobatVersion
        DotNet35 = $DotNet35; Dedalus = $DedalusInstalado; ESET = $ESETInstalado; WinRAR = $WinRARInstalado
    }
}

function Send-SoftwareInfo {
    param([string]$InvST)

    Write-StepHeader -Step 25 -Title "REGISTRANDO INVENTARIO DE SOFTWARE"
    Show-ProgressCosmos -Step 25
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $Info = Get-SoftwareInfo
    $Body = @{
        Accion = "software"; InvST = $InvST; NombreEquipo = $Info.NombreEquipo
        WindowsVersion = $Info.WindowsVersion; WindowsBuild = $Info.WindowsBuild
        WindowsActivado = $Info.WindowsActivado; ProductKey = $Info.ProductKey
        Office = $Info.Office; Chrome = $Info.Chrome; Acrobat = $Info.Acrobat
        DotNet35 = $Info.DotNet35; Dedalus = $Info.Dedalus; ESET = $Info.ESET
        WinRAR = $Info.WinRAR; UsuarioWindows = $Info.UsuarioWindows
        FechaConfig = (Get-Date -Format "dd/MM/yyyy HH:mm")
    } | ConvertTo-Json

    try {
        $Response = Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 60
        if ($Response.status -eq "OK") {
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |   INVENTARIO SOFTWARE REGISTRADO         |" -ForegroundColor Green
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |  Windows:  $($Info.WindowsVersion) (Build $($Info.WindowsBuild))" -ForegroundColor White
            Write-Host "  |  Activado: $($Info.WindowsActivado)" -ForegroundColor $(if ($Info.WindowsActivado -eq "Si") { "Green" } else { "Red" })
            Write-Host "  |  Office:   $($Info.Office)" -ForegroundColor White
            Write-Host "  |  Chrome:   $($Info.Chrome)" -ForegroundColor White
            Write-Host "  |  ESET:     $($Info.ESET)" -ForegroundColor $(if ($Info.ESET -eq "Si") { "Green" } else { "Red" })
            Write-Host "  |  WinRAR:   $($Info.WinRAR)" -ForegroundColor White
            Write-Host "  |  Dedalus:  $($Info.Dedalus)" -ForegroundColor White
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            return $true
        }
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
    return $false
}

# --- FUNCIONES DE CONFIGURACION ---

function Connect-Servidor {
    Write-StepHeader -Step 2 -Title "CONECTANDO AL SERVIDOR"
    Show-ProgressCosmos -Step 2
    net use \\$Servidor /delete /y 2>$null | Out-Null
    $Result = net use \\$Servidor\soportefaa /user:$Usuario $Password /persistent:no 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Log "Conectado a \\$Servidor\soportefaa" "OK"
        return $true
    } else {
        Write-Log "Error al conectar: $Result" "ERROR"
        return $false
    }
}

function Remove-OfficePrevio {
    Write-StepHeader -Step 3 -Title "ELIMINANDO VERSIONES PREVIAS DE OFFICE"
    Show-ProgressCosmos -Step 3

    $Eliminado = $false

    # --- 1. Desinstalar Office Click-to-Run (Microsoft 365, OneNote, etc.) ---
    $CTR = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe"
    if (Test-Path $CTR) {
        $Config = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
        if ($Config -and $Config.ProductReleaseIds) {
            $ProductIds = $Config.ProductReleaseIds
            Write-Log "Office Click-to-Run detectado: $ProductIds"

            # Desinstalar cada producto por separado
            $Productos = $ProductIds -split ","
            foreach ($prod in $Productos) {
                $prod = $prod.Trim()
                if ($prod) {
                    Write-Log "Removiendo: $prod"
                    Start-Process -FilePath $CTR -ArgumentList "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=$prod DisplayLevel=False" -Wait -NoNewWindow -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 3
                }
            }

            # Esperar a que termine el proceso de desinstalacion
            $Intentos = 0
            while ((Get-Process -Name "OfficeClickToRun" -ErrorAction SilentlyContinue) -and $Intentos -lt 30) {
                Start-Sleep -Seconds 2
                $Intentos++
            }

            Write-Log "Office Click-to-Run desinstalado" "OK"
            $Eliminado = $true
        }
    }

    # --- 2. Desinstalar por registro (captura cualquier version restante) ---
    $RegPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($rp in $RegPaths) {
        Get-ItemProperty $rp -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -and (
                $_.DisplayName -like "*Microsoft 365*" -or
                $_.DisplayName -like "*Office 365*" -or
                $_.DisplayName -like "*OneNote*es-es*"
            )
        } | ForEach-Object {
            $Nombre = $_.DisplayName
            $Uninstall = $_.UninstallString
            if ($Uninstall) {
                Write-Log "Desinstalando: $Nombre"
                try {
                    # Usar cmd /c para ejecutar la cadena de desinstalacion tal cual
                    Start-Process "cmd.exe" -ArgumentList "/c $Uninstall DisplayLevel=False" -Wait -NoNewWindow -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 3
                    Write-Log "Desinstalado: $Nombre" "OK"
                } catch {
                    Write-Log "Error al desinstalar $Nombre : $($_.Exception.Message)" "WARN"
                }
                $Eliminado = $true
            }
        }
    }

    # --- 3. Verificar que se elimino ---
    Start-Sleep -Seconds 5
    $Restante = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*Microsoft 365*" -or $_.DisplayName -like "*Office 365*" }

    if ($Restante) {
        Write-Log "Aun quedan restos de Office, intentando limpieza final..." "WARN"
        foreach ($r in $Restante) {
            if ($r.UninstallString) {
                Start-Process "cmd.exe" -ArgumentList "/c $($r.UninstallString) DisplayLevel=False" -Wait -NoNewWindow -ErrorAction SilentlyContinue
            }
        }
    }

    if ($Eliminado) {
        $Script:SoftwareInstalado += "Office previo removido"
    } else {
        Write-Log "No se encontraron versiones previas de Office" "INFO"
    }
}

function New-UsuarioSoporte {
    Write-StepHeader -Step 4 -Title "CREANDO USUARIO SOPORTE (ADMINISTRADOR)"
    Show-ProgressCosmos -Step 4

    $Existe = Get-LocalUser -Name $UsuarioSoporte -ErrorAction SilentlyContinue
    if (-not $Existe) {
        $SecurePass = ConvertTo-SecureString $PasswordSoporte -AsPlainText -Force
        New-LocalUser -Name $UsuarioSoporte -Password $SecurePass -Description "Soporte Tecnico HCG" -PasswordNeverExpires -UserMayNotChangePassword | Out-Null
        Write-Log "Usuario '$UsuarioSoporte' creado" "OK"
    } else {
        Write-Log "Usuario '$UsuarioSoporte' ya existe" "INFO"
    }

    # Asegurar que es administrador
    Add-LocalGroupMember -Group "Administradores" -Member $UsuarioSoporte -ErrorAction SilentlyContinue
    Write-Log "Usuario '$UsuarioSoporte' es Administrador" "OK"
    $Script:SoftwareInstalado += "Usuario Soporte"
}

function New-UsuarioEquipo {
    param([string]$NumInventario)

    Write-StepHeader -Step 5 -Title "CREANDO USUARIO NORMAL ($NumInventario)"
    Show-ProgressCosmos -Step 5

    $NombreUsuario = $NumInventario
    $Existe = Get-LocalUser -Name $NombreUsuario -ErrorAction SilentlyContinue

    if (-not $Existe) {
        New-LocalUser -Name $NombreUsuario -NoPassword -Description "Usuario Equipo $NumInventario - HCG FAA" -PasswordNeverExpires -UserMayNotChangePassword | Out-Null
        Write-Log "Usuario '$NombreUsuario' creado" "OK"
    } else {
        Write-Log "Usuario '$NombreUsuario' ya existe" "INFO"
    }

    # Asegurar que es usuario estandar (grupo Usuarios)
    Add-LocalGroupMember -Group "Usuarios" -Member $NombreUsuario -ErrorAction SilentlyContinue

    # Asegurar que NO es administrador
    Remove-LocalGroupMember -Group "Administradores" -Member $NombreUsuario -ErrorAction SilentlyContinue

    # Habilitar el usuario
    Enable-LocalUser -Name $NombreUsuario -ErrorAction SilentlyContinue

    # Configurar auto-login para que este usuario inicie sesion automaticamente
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    Set-ItemProperty -Path $RegPath -Name "AutoAdminLogon" -Value "1"
    Set-ItemProperty -Path $RegPath -Name "DefaultUserName" -Value $NombreUsuario
    Set-ItemProperty -Path $RegPath -Name "DefaultPassword" -Value ""
    Set-ItemProperty -Path $RegPath -Name "DefaultDomainName" -Value $env:COMPUTERNAME
    Write-Log "Auto-login configurado para '$NombreUsuario'" "OK"

    # Crear script de primer inicio para configurar barra de tareas del nuevo usuario
    $FirstLoginScript = @'
# HCG - Configurar barra de tareas en primer inicio de sesion
Start-Sleep -Seconds 15
$TaskBarFolder = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
if (-not (Test-Path $TaskBarFolder)) { New-Item -ItemType Directory -Path $TaskBarFolder -Force | Out-Null }
$ChromeShortcut = "C:\Users\Public\Desktop\Google Chrome.lnk"
if (Test-Path $ChromeShortcut) { Copy-Item -Path $ChromeShortcut -Destination $TaskBarFolder -Force }
$xHISShortcut = "C:\Dedalus\xHIS\xHIS v6.lnk"
if (Test-Path $xHISShortcut) { Copy-Item -Path $xHISShortcut -Destination $TaskBarFolder -Force }
foreach ($folder in @("C:\Dedalus\EscritorioClinico", "C:\Dedalus\xFARMA", "C:\Dedalus\hPRESC")) {
    if (Test-Path $folder) {
        $lnk = Get-ChildItem -Path $folder -Filter "*.lnk" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($lnk) { Copy-Item -Path $lnk.FullName -Destination $TaskBarFolder -Force }
    }
}
'@
    if (-not (Test-Path "C:\HCG_Logs")) { New-Item -ItemType Directory -Path "C:\HCG_Logs" -Force | Out-Null }
    $FirstLoginScript | Out-File -FilePath "C:\HCG_Logs\setup_firstlogin.ps1" -Encoding UTF8 -Force

    # Configurar perfil por defecto para heredar configuracion al nuevo usuario
    try {
        reg load "HKU\DefaultUser" "C:\Users\Default\NTUSER.DAT" 2>$null
        $DefRunOnce = "Registry::HKU\DefaultUser\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
        if (-not (Test-Path $DefRunOnce)) { New-Item -Path $DefRunOnce -Force | Out-Null }
        Set-ItemProperty -Path $DefRunOnce -Name "HCG_Setup" -Value 'powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\HCG_Logs\setup_firstlogin.ps1"'
        [gc]::Collect()
        reg unload "HKU\DefaultUser" 2>$null
        Write-Log "Script de primer inicio configurado" "OK"
    } catch {
        [gc]::Collect()
        reg unload "HKU\DefaultUser" 2>$null
        Write-Log "No se pudo configurar script de primer inicio" "WARN"
    }

    Write-Log "Usuario '$NombreUsuario' listo (estandar, auto-login)" "OK"
    $Script:SoftwareInstalado += "Usuario $NombreUsuario"
}

function Set-ImagenesUsuarios {
    param([string]$NumInventario)

    Write-StepHeader -Step 6 -Title "CONFIGURANDO IMAGENES DE PERFIL"
    Show-ProgressCosmos -Step 6

    try {
        Add-Type -AssemblyName System.Drawing

        $Size = 448

        # =================================================================
        # PALETA DE COLORES - Tonos suaves, relajantes y profesionales
        # Cada equipo recibe una combinacion unica basada en su inventario
        # =================================================================
        $PaletaColores = @(
            @{ Fondo1 = @(100, 149, 237); Fondo2 = @(65, 105, 180);  Nombre = "Cielo" },
            @{ Fondo1 = @(72, 191, 170);  Fondo2 = @(45, 150, 140);  Nombre = "Jade" },
            @{ Fondo1 = @(180, 136, 200); Fondo2 = @(140, 100, 170); Nombre = "Lavanda" },
            @{ Fondo1 = @(210, 150, 100); Fondo2 = @(175, 120, 75);  Nombre = "Ambar" },
            @{ Fondo1 = @(95, 170, 200);  Fondo2 = @(60, 130, 165);  Nombre = "Oceano" },
            @{ Fondo1 = @(120, 190, 130); Fondo2 = @(80, 155, 95);   Nombre = "Esmeralda" },
            @{ Fondo1 = @(200, 130, 150); Fondo2 = @(165, 95, 120);  Nombre = "Coral" },
            @{ Fondo1 = @(150, 160, 200); Fondo2 = @(110, 120, 170); Nombre = "Lila" },
            @{ Fondo1 = @(200, 175, 110); Fondo2 = @(170, 145, 80);  Nombre = "Miel" },
            @{ Fondo1 = @(110, 180, 190); Fondo2 = @(75, 145, 160);  Nombre = "Turquesa" },
            @{ Fondo1 = @(175, 140, 170); Fondo2 = @(140, 105, 140); Nombre = "Amatista" },
            @{ Fondo1 = @(140, 185, 140); Fondo2 = @(100, 150, 105); Nombre = "Salvia" },
            @{ Fondo1 = @(190, 155, 130); Fondo2 = @(155, 120, 95);  Nombre = "Arena" },
            @{ Fondo1 = @(130, 165, 210); Fondo2 = @(90, 130, 180);  Nombre = "Zafiro" },
            @{ Fondo1 = @(185, 160, 185); Fondo2 = @(150, 125, 150); Nombre = "Orquidea" },
            @{ Fondo1 = @(160, 195, 165); Fondo2 = @(120, 160, 130); Nombre = "Menta" }
        )

        # Seleccionar color basado en el numero de inventario (determinista)
        $Seed = [int]($NumInventario) % $PaletaColores.Count
        $Colores = $PaletaColores[$Seed]

        # =================================================================
        # AVATAR SOPORTE: Elegante, profesional con "S"
        # =================================================================
        $BmpS = New-Object System.Drawing.Bitmap $Size, $Size
        $GfxS = [System.Drawing.Graphics]::FromImage($BmpS)
        $GfxS.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $GfxS.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

        # Fondo gradiente gris oscuro elegante
        $RectS = New-Object System.Drawing.Rectangle(0, 0, $Size, $Size)
        $GradS = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $RectS,
            [System.Drawing.Color]::FromArgb(55, 60, 70),
            [System.Drawing.Color]::FromArgb(35, 40, 50),
            [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal
        )
        $GfxS.FillRectangle($GradS, $RectS)

        # Circulo principal gris azulado
        $BrushCircle = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(75, 85, 100))
        $GfxS.FillEllipse($BrushCircle, 30, 30, ($Size - 60), ($Size - 60))

        # Anillo sutil
        $PenRing = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(100, 200, 200, 210)), 3
        $GfxS.DrawEllipse($PenRing, 30, 30, ($Size - 60), ($Size - 60))

        # Letra S
        $SFS = New-Object System.Drawing.StringFormat
        $SFS.Alignment = [System.Drawing.StringAlignment]::Center
        $SFS.LineAlignment = [System.Drawing.StringAlignment]::Center
        $FontS = New-Object System.Drawing.Font("Segoe UI Light", 140)
        $BrushText = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(220, 225, 230))
        $GfxS.DrawString("S", $FontS, $BrushText, (New-Object System.Drawing.RectangleF(0, 0, $Size, $Size)), $SFS)

        $AvatarSoporte = "$env:TEMP\avatar_soporte.jpg"
        $BmpS.Save($AvatarSoporte, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $FontS.Dispose(); $BrushText.Dispose(); $BrushCircle.Dispose()
        $PenRing.Dispose(); $GradS.Dispose(); $SFS.Dispose()
        $GfxS.Dispose(); $BmpS.Dispose()

        # =================================================================
        # AVATAR USUARIO: Minimalista, suave y unico por equipo
        # =================================================================
        $AvatarUsuario = "$env:TEMP\avatar_usuario.jpg"

        $BmpU = New-Object System.Drawing.Bitmap $Size, $Size
        $GfxU = [System.Drawing.Graphics]::FromImage($BmpU)
        $GfxU.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $GfxU.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

        # Fondo gradiente suave con los colores del equipo
        $RectU = New-Object System.Drawing.Rectangle(0, 0, $Size, $Size)
        $Color1 = [System.Drawing.Color]::FromArgb($Colores.Fondo1[0], $Colores.Fondo1[1], $Colores.Fondo1[2])
        $Color2 = [System.Drawing.Color]::FromArgb($Colores.Fondo2[0], $Colores.Fondo2[1], $Colores.Fondo2[2])

        # Angulo del gradiente segun inventario
        $AnguloGrad = @(
            [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal,
            [System.Drawing.Drawing2D.LinearGradientMode]::BackwardDiagonal,
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical,
            [System.Drawing.Drawing2D.LinearGradientMode]::Horizontal
        )
        $GradMode = $AnguloGrad[[int]($NumInventario) % $AnguloGrad.Count]
        $GradU = New-Object System.Drawing.Drawing2D.LinearGradientBrush($RectU, $Color1, $Color2, $GradMode)
        $GfxU.FillRectangle($GradU, $RectU)

        # Efecto de luz suave (circulo semi-transparente arriba-izquierda)
        $BrushGlow = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(25, 255, 255, 255))
        $GfxU.FillEllipse($BrushGlow, -100, -120, ($Size + 80), ($Size))

        # Anillo circular exterior sutil
        $PenMarco = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(60, 255, 255, 255)), 3
        $GfxU.DrawEllipse($PenMarco, 20, 20, ($Size - 40), ($Size - 40))

        # Numero de inventario centrado
        $SFU = New-Object System.Drawing.StringFormat
        $SFU.Alignment = [System.Drawing.StringAlignment]::Center
        $SFU.LineAlignment = [System.Drawing.StringAlignment]::Center

        $FontNum = New-Object System.Drawing.Font("Segoe UI Light", 90)
        # Sombra suave
        $BrushSombra = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(50, 0, 0, 0))
        $GfxU.DrawString($NumInventario, $FontNum, $BrushSombra, (New-Object System.Drawing.RectangleF(3, 3, $Size, $Size)), $SFU)
        # Texto blanco
        $BrushBlanco = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(240, 255, 255, 255))
        $GfxU.DrawString($NumInventario, $FontNum, $BrushBlanco, (New-Object System.Drawing.RectangleF(0, 0, $Size, $Size)), $SFU)

        # Texto "HCG" abajo, pequeno y discreto
        $FontHCG = New-Object System.Drawing.Font("Segoe UI Light", 22)
        $BrushHCG = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(140, 255, 255, 255))
        $GfxU.DrawString("HCG", $FontHCG, $BrushHCG, (New-Object System.Drawing.RectangleF(0, ($Size - 100), $Size, 70)), $SFU)

        # Guardar
        $BmpU.Save($AvatarUsuario, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        Write-Log "Avatar personalizado: tono $($Colores.Nombre)" "OK"

        # Limpiar recursos
        $FontNum.Dispose(); $FontHCG.Dispose()
        $BrushBlanco.Dispose(); $BrushSombra.Dispose(); $BrushHCG.Dispose()
        $BrushGlow.Dispose(); $PenMarco.Dispose(); $GradU.Dispose(); $SFU.Dispose()
        $GfxU.Dispose(); $BmpU.Dispose()

        # =================================================================
        # APLICAR AVATARES CON API DE WINDOWS
        # =================================================================
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class UserTileAPI {
    [DllImport("shell32.dll", EntryPoint = "#262", CharSet = CharSet.Unicode, PreserveSig = false)]
    public static extern void SetUserTile(string username, int reserved, string imagePath);
}
"@ -ErrorAction SilentlyContinue

        [UserTileAPI]::SetUserTile($UsuarioSoporte, 0, $AvatarSoporte)
        Write-Log "Imagen de perfil: '$UsuarioSoporte' configurada" "OK"

        [UserTileAPI]::SetUserTile($NumInventario, 0, $AvatarUsuario)
        Write-Log "Imagen de perfil: '$NumInventario' -> $($Colores.Nombre)" "OK"

    } catch {
        Write-Log "No se pudieron configurar imagenes: $($_.Exception.Message)" "WARN"
    }
}

function Set-RedPrivada {
    Write-StepHeader -Step 7 -Title "CONFIGURANDO RED PRIVADA"
    Show-ProgressCosmos -Step 7
    Get-NetConnectionProfile | ForEach-Object { Set-NetConnectionProfile -InterfaceIndex $_.InterfaceIndex -NetworkCategory Private -ErrorAction SilentlyContinue }
    netsh advfirewall firewall set rule group="Network Discovery" new enable=Yes 2>$null
    netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes 2>$null
    Write-Log "Red configurada como privada" "OK"
    $Script:SoftwareInstalado += "Red privada"
}

function Set-HoraAutomatica {
    Write-StepHeader -Step 8 -Title "CONFIGURANDO HORA AUTOMATICA"
    Show-ProgressCosmos -Step 8
    Set-TimeZone -Id $ZonaHoraria -ErrorAction SilentlyContinue
    Set-Service -Name w32time -StartupType Automatic
    Start-Service w32time -ErrorAction SilentlyContinue
    w32tm /config /manualpeerlist:"time.windows.com" /syncfromflags:manual /reliable:yes /update 2>$null
    w32tm /resync /force 2>$null
    Write-Log "Hora automatica configurada (Guadalajara)" "OK"
    $Script:SoftwareInstalado += "Hora auto"
}

function Set-TemaOscuro {
    Write-StepHeader -Step 9 -Title "CONFIGURANDO TEMA OSCURO PERSONALIZADO"
    Show-ProgressCosmos -Step 9

    # =================================================================
    # PALETA DE COLORES DE ACENTO (misma que avatares - suave y relajante)
    # Windows usa formato ABGR (invertido) para AccentColor en el registro
    # =================================================================
    $PaletaAcento = @(
        @{ R = 100; G = 149; B = 237; Nombre = "Cielo" },
        @{ R = 72;  G = 191; B = 170; Nombre = "Jade" },
        @{ R = 180; G = 136; B = 200; Nombre = "Lavanda" },
        @{ R = 210; G = 150; B = 100; Nombre = "Ambar" },
        @{ R = 95;  G = 170; B = 200; Nombre = "Oceano" },
        @{ R = 120; G = 190; B = 130; Nombre = "Esmeralda" },
        @{ R = 200; G = 130; B = 150; Nombre = "Coral" },
        @{ R = 150; G = 160; B = 200; Nombre = "Lila" },
        @{ R = 200; G = 175; B = 110; Nombre = "Miel" },
        @{ R = 110; G = 180; B = 190; Nombre = "Turquesa" },
        @{ R = 175; G = 140; B = 170; Nombre = "Amatista" },
        @{ R = 140; G = 185; B = 140; Nombre = "Salvia" },
        @{ R = 190; G = 155; B = 130; Nombre = "Arena" },
        @{ R = 130; G = 165; B = 210; Nombre = "Zafiro" },
        @{ R = 185; G = 160; B = 185; Nombre = "Orquidea" },
        @{ R = 160; G = 195; B = 165; Nombre = "Menta" }
    )

    # Seleccionar color basado en el numero de inventario (mismo que avatar)
    $Seed = [int]($NumInventario) % $PaletaAcento.Count
    $ColorAccent = $PaletaAcento[$Seed]

    # Windows AccentColor usa formato ABGR: 0xFF + BB + GG + RR
    $HexAccent = "0xFF{0:X2}{1:X2}{2:X2}" -f $ColorAccent.B, $ColorAccent.G, $ColorAccent.R
    $AccentABGR = [Convert]::ToInt64($HexAccent, 16)
    # Version mas oscura para elementos inactivos
    $R2 = [Math]::Max(0, $ColorAccent.R - 30)
    $G2 = [Math]::Max(0, $ColorAccent.G - 30)
    $B2 = [Math]::Max(0, $ColorAccent.B - 30)
    $HexMenu = "0xFF{0:X2}{1:X2}{2:X2}" -f $B2, $G2, $R2
    $AccentMenuABGR = [Convert]::ToInt64($HexMenu, 16)

    $RegPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    Set-ItemProperty -Path $RegPath -Name "AppsUseLightTheme" -Value 0 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $RegPath -Name "SystemUsesLightTheme" -Value 0 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $RegPath -Name "ColorPrevalence" -Value 1 -Type DWord -ErrorAction SilentlyContinue

    # Color de acento en barra de titulo y bordes
    $DWMPath = "HKCU:\SOFTWARE\Microsoft\Windows\DWM"
    Set-ItemProperty -Path $DWMPath -Name "ColorPrevalence" -Value 1 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $DWMPath -Name "AccentColor" -Value $AccentABGR -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $DWMPath -Name "AccentColorInactive" -Value $AccentMenuABGR -Type DWord -ErrorAction SilentlyContinue

    # Color de acento en Start menu y barra de tareas
    $ExplorerAccent = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent"
    if (-not (Test-Path $ExplorerAccent)) { New-Item -Path $ExplorerAccent -Force | Out-Null }
    Set-ItemProperty -Path $ExplorerAccent -Name "AccentColorMenu" -Value $AccentMenuABGR -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $ExplorerAccent -Name "StartColorMenu" -Value $AccentABGR -Type DWord -ErrorAction SilentlyContinue

    Write-Log "Color de acento: $($ColorAccent.Nombre)" "OK"

    # Aplicar tambien al perfil por defecto (para el usuario de inventario)
    try {
        reg load "HKU\DefaultUser" "C:\Users\Default\NTUSER.DAT" 2>$null
        $DefPath = "Registry::HKU\DefaultUser\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        if (-not (Test-Path $DefPath)) { New-Item -Path $DefPath -Force | Out-Null }
        Set-ItemProperty -Path $DefPath -Name "AppsUseLightTheme" -Value 0 -Type DWord
        Set-ItemProperty -Path $DefPath -Name "SystemUsesLightTheme" -Value 0 -Type DWord
        Set-ItemProperty -Path $DefPath -Name "ColorPrevalence" -Value 1 -Type DWord

        $DefDWM = "Registry::HKU\DefaultUser\SOFTWARE\Microsoft\Windows\DWM"
        if (-not (Test-Path $DefDWM)) { New-Item -Path $DefDWM -Force | Out-Null }
        Set-ItemProperty -Path $DefDWM -Name "ColorPrevalence" -Value 1 -Type DWord
        Set-ItemProperty -Path $DefDWM -Name "AccentColor" -Value $AccentABGR -Type DWord
        Set-ItemProperty -Path $DefDWM -Name "AccentColorInactive" -Value $AccentMenuABGR -Type DWord

        $DefAccent = "Registry::HKU\DefaultUser\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent"
        if (-not (Test-Path $DefAccent)) { New-Item -Path $DefAccent -Force | Out-Null }
        Set-ItemProperty -Path $DefAccent -Name "AccentColorMenu" -Value $AccentMenuABGR -Type DWord
        Set-ItemProperty -Path $DefAccent -Name "StartColorMenu" -Value $AccentABGR -Type DWord

        [gc]::Collect()
        reg unload "HKU\DefaultUser" 2>$null
        Write-Log "Tema personalizado aplicado al perfil por defecto" "OK"
    } catch {
        [gc]::Collect()
        reg unload "HKU\DefaultUser" 2>$null
    }

    Write-Log "Tema oscuro personalizado: $($ColorAccent.Nombre)" "OK"
    $Script:SoftwareInstalado += "Tema personalizado"
}

function Set-FondoPantalla {
    Write-StepHeader -Step 14 -Title "ESTABLECIENDO FONDO DE PANTALLA"
    Show-ProgressCosmos -Step 14

    $FondoLocal = "C:\Windows\Web\Wallpaper\HCG_Fondo.png"
    $NombrePC = "PC-$NumInventario"
    $FechaConfig = Get-Date -Format "dd/MM/yyyy"

    # =================================================================
    # METODO 1: HTML + Chrome headless (diseño completo CSS)
    # =================================================================
    $TemplateHTML = "$RutaWallpaper\wallpaper_template.html"
    $ImagenFondo = "$RutaWallpaper\wallpaper_hcg.jpg"
    $ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    $WallpaperGenerado = $false

    if ((Test-Path $TemplateHTML) -and (Test-Path $ImagenFondo) -and (Test-Path $ChromePath)) {
        Write-Log "Generando wallpaper con HTML + Chrome..."

        try {
            # Crear carpeta temporal para el renderizado
            $TempDir = "$env:TEMP\HCG_Wallpaper"
            if (-not (Test-Path $TempDir)) { New-Item -ItemType Directory -Path $TempDir -Force | Out-Null }

            # Copiar imagen de fondo a carpeta temporal
            Copy-Item -Path $ImagenFondo -Destination "$TempDir\wallpaper_hcg.jpg" -Force

            # Leer template y reemplazar variables
            $HTMLContent = Get-Content -Path $TemplateHTML -Raw -Encoding UTF8
            $HTMLContent = $HTMLContent -replace '\{\{HOSTNAME\}\}', $NombrePC
            $HTMLContent = $HTMLContent -replace '\{\{INVENTARIO\}\}', $NumInventario
            $HTMLContent = $HTMLContent -replace '\{\{FECHA\}\}', $FechaConfig
            $HTMLContent = $HTMLContent -replace '\{\{EXT\}\}', '54425'

            # Guardar HTML personalizado
            $HTMLLocal = "$TempDir\wallpaper_render.html"
            $HTMLContent | Out-File -FilePath $HTMLLocal -Encoding UTF8 -Force

            # Renderizar con Chrome headless a PNG
            $ScreenshotPath = "$TempDir\wallpaper_screenshot.png"
            $FileUrl = "file:///$($HTMLLocal -replace '\\','/')"

            $ChromeArgs = @(
                "--headless=new"
                "--disable-gpu"
                "--no-sandbox"
                "--disable-software-rasterizer"
                "--window-size=1920,1080"
                "--screenshot=`"$ScreenshotPath`""
                "--hide-scrollbars"
                "--default-background-color=00000000"
                "`"$FileUrl`""
            )

            $ChromeProcess = Start-Process -FilePath $ChromePath -ArgumentList ($ChromeArgs -join " ") -Wait -PassThru -NoNewWindow -ErrorAction Stop

            # Esperar a que se genere el archivo
            Start-Sleep -Seconds 2

            if (Test-Path $ScreenshotPath) {
                Copy-Item -Path $ScreenshotPath -Destination $FondoLocal -Force
                Write-Log "Wallpaper HTML generado con Chrome headless" "OK"
                $WallpaperGenerado = $true
            } else {
                Write-Log "Chrome no genero el screenshot, usando fallback" "WARN"
            }

            # Limpiar temporales
            Remove-Item -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Log "Error al generar wallpaper HTML: $($_.Exception.Message)" "WARN"
        }
    } else {
        if (-not (Test-Path $TemplateHTML)) { Write-Log "Template HTML no encontrado, usando fallback" "WARN" }
        if (-not (Test-Path $ChromePath)) { Write-Log "Chrome no instalado aun, usando fallback" "WARN" }
    }

    # =================================================================
    # METODO 2 (FALLBACK): System.Drawing sobre imagen JPG
    # =================================================================
    if (-not $WallpaperGenerado) {
        $Fondo = "$RutaWallpaper\wallpaper_hcg.jpg"
        if (-not (Test-Path $Fondo)) { $Fondo = "$RutaAccesos\Fondo de Pantalla.jpg" }
        if (-not (Test-Path $Fondo)) { $Fondo = Find-Installer -Carpeta $RutaWallpaper -Filtro "*.jpg" }

        $FondoLocal = "C:\Windows\Web\Wallpaper\HCG_Fondo.jpg"

        if ($Fondo -and (Test-Path $Fondo)) {
            try {
                Add-Type -AssemblyName System.Drawing

                $ImgOrig = [System.Drawing.Image]::FromFile($Fondo)
                $Bmp = New-Object System.Drawing.Bitmap($ImgOrig.Width, $ImgOrig.Height)
                $Gfx = [System.Drawing.Graphics]::FromImage($Bmp)
                $Gfx.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $Gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $Gfx.DrawImage($ImgOrig, 0, 0, $ImgOrig.Width, $ImgOrig.Height)

                $W = $ImgOrig.Width; $H = $ImgOrig.Height

                # Franja inferior
                $AlturaFranja = 52
                $BrushFranja = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(140, 0, 0, 0))
                $Gfx.FillRectangle($BrushFranja, 0, ($H - $AlturaFranja), $W, $AlturaFranja)

                $PenLinea = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(60, 255, 255, 255)), 1
                $Gfx.DrawLine($PenLinea, 0, ($H - $AlturaFranja), $W, ($H - $AlturaFranja))

                $SF = New-Object System.Drawing.StringFormat
                $SF.LineAlignment = [System.Drawing.StringAlignment]::Center
                $FontInfo = New-Object System.Drawing.Font("Segoe UI", 13)
                $BrushWhite = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(220, 255, 255, 255))
                $BrushGray = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(160, 255, 255, 255))
                $YFranja = $H - $AlturaFranja

                $SF.Alignment = [System.Drawing.StringAlignment]::Near
                $TextoIzq = "  $NombrePC  |  Inv. ST: $NumInventario  |  Configurado: $FechaConfig"
                $RectIzq = New-Object System.Drawing.RectangleF(10, $YFranja, ($W / 2), $AlturaFranja)
                $Gfx.DrawString($TextoIzq, $FontInfo, $BrushWhite, $RectIzq, $SF)

                $SF.Alignment = [System.Drawing.StringAlignment]::Far
                $TextoDer = "Soporte Tecnico - Ext. 54425  |  Hospital Civil FAA  "
                $RectDer = New-Object System.Drawing.RectangleF(($W / 2), $YFranja, (($W / 2) - 10), $AlturaFranja)
                $Gfx.DrawString($TextoDer, $FontInfo, $BrushGray, $RectDer, $SF)

                $Bmp.Save($FondoLocal, [System.Drawing.Imaging.ImageFormat]::Jpeg)
                $FontInfo.Dispose(); $BrushWhite.Dispose(); $BrushGray.Dispose()
                $BrushFranja.Dispose(); $PenLinea.Dispose(); $SF.Dispose()
                $Gfx.Dispose(); $Bmp.Dispose(); $ImgOrig.Dispose()

                Write-Log "Wallpaper fallback con System.Drawing" "OK"
                $WallpaperGenerado = $true
            } catch {
                Copy-Item -Path $Fondo -Destination $FondoLocal -Force
                Write-Log "Wallpaper copiado sin personalizar: $($_.Exception.Message)" "WARN"
                $WallpaperGenerado = $true
            }
        }
    }

    # =================================================================
    # APLICAR WALLPAPER AL SISTEMA
    # =================================================================
    if ($WallpaperGenerado -and (Test-Path $FondoLocal)) {
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper" -Value $FondoLocal
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperStyle" -Value "10"
        rundll32.exe user32.dll, UpdatePerUserSystemParameters 1, True

        # Aplicar tambien al perfil por defecto (para el usuario de inventario)
        try {
            reg load "HKU\DefaultUser" "C:\Users\Default\NTUSER.DAT" 2>$null
            $DefDesktop = "Registry::HKU\DefaultUser\Control Panel\Desktop"
            if (-not (Test-Path $DefDesktop)) { New-Item -Path $DefDesktop -Force | Out-Null }
            Set-ItemProperty -Path $DefDesktop -Name "Wallpaper" -Value $FondoLocal
            Set-ItemProperty -Path $DefDesktop -Name "WallpaperStyle" -Value "10"
            [gc]::Collect()
            reg unload "HKU\DefaultUser" 2>$null
            Write-Log "Fondo aplicado al perfil por defecto" "OK"
        } catch {
            [gc]::Collect()
            reg unload "HKU\DefaultUser" 2>$null
        }

        Write-Log "Fondo de pantalla establecido" "OK"
        $Script:SoftwareInstalado += "Fondo HCG"
    } else {
        Write-Log "No se encontro imagen de fondo" "WARN"
    }
}

function Install-WinRAR {
    Write-StepHeader -Step 10 -Title "INSTALANDO WINRAR CON LICENCIA"
    Show-ProgressCosmos -Step 10

    # PASO 1: Instalar WinRAR
    $YaInstalado = Test-Path "C:\Program Files\WinRAR\WinRAR.exe"

    if (-not $YaInstalado) {
        # Buscar el instalador (winrar-x64-*.exe)
        $Instalador = "$RutaWinRAR\winrar-x64-713.exe"
        if (-not (Test-Path $Instalador)) {
            # Buscar cualquier instalador winrar-x64
            $Instalador = Get-ChildItem -Path $RutaWinRAR -Filter "winrar-x64*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($Instalador) { $Instalador = $Instalador.FullName }
        }
        if (-not $Instalador -or -not (Test-Path $Instalador)) {
            # Buscar cualquier exe que NO sea la licencia
            $Instalador = Get-ChildItem -Path $RutaWinRAR -Filter "*.exe" -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*License*" } | Select-Object -First 1
            if ($Instalador) { $Instalador = $Instalador.FullName }
        }

        if ($Instalador -and (Test-Path $Instalador)) {
            Write-Log "Instalando WinRAR: $(Split-Path $Instalador -Leaf)"
            Start-Process -FilePath $Instalador -ArgumentList "/S" -Wait -NoNewWindow
            Start-Sleep -Seconds 3

            if (Test-Path "C:\Program Files\WinRAR\WinRAR.exe") {
                Write-Log "WinRAR instalado correctamente" "OK"
            } else {
                Write-Log "No se pudo verificar la instalacion de WinRAR" "WARN"
            }
        } else {
            Write-Log "No se encontro instalador de WinRAR en: $RutaWinRAR" "ERROR"
            return
        }
    } else {
        Write-Log "WinRAR ya esta instalado" "INFO"
    }

    # PASO 2: Aplicar licencia
    $Licencia = "$RutaWinRAR\WinRAR License.exe"
    if (-not (Test-Path $Licencia)) {
        $Licencia = Get-ChildItem -Path $RutaWinRAR -Filter "*License*" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($Licencia) { $Licencia = $Licencia.FullName }
    }

    if ($Licencia -and (Test-Path $Licencia)) {
        Write-Log "Aplicando licencia de WinRAR..."
        Start-Process -FilePath $Licencia -ArgumentList "/S" -Wait -NoNewWindow
        Start-Sleep -Seconds 2

        # Verificar que la licencia se aplico (buscar rarreg.key en la carpeta de WinRAR)
        if (Test-Path "C:\Program Files\WinRAR\rarreg.key") {
            Write-Log "Licencia de WinRAR aplicada correctamente" "OK"
        } else {
            Write-Log "Licencia ejecutada (verificar activacion manualmente)" "WARN"
        }
    } else {
        Write-Log "No se encontro archivo de licencia de WinRAR" "WARN"
    }

    $Script:SoftwareInstalado += "WinRAR"
}

function Install-DotNet35 {
    Write-StepHeader -Step 11 -Title "INSTALANDO .NET FRAMEWORK 3.5"
    Show-ProgressCosmos -Step 11

    $Estado = (Get-WindowsOptionalFeature -Online -FeatureName "NetFx3" -ErrorAction SilentlyContinue).State
    if ($Estado -eq "Enabled") {
        Write-Log ".NET 3.5 ya instalado" "INFO"
        $Script:SoftwareInstalado += ".NET 3.5"
        return
    }

    # Intentar instalacion offline desde el servidor (sin descargar de internet)
    if (Test-Path $RutaDotNet) {
        Write-Log "Instalando .NET 3.5 offline desde servidor..."
        try {
            $DismResult = Dism /Online /Enable-Feature /FeatureName:NetFx3 /All /Source:"$RutaDotNet" /LimitAccess /NoRestart 2>&1
            $EstadoPost = (Get-WindowsOptionalFeature -Online -FeatureName "NetFx3" -ErrorAction SilentlyContinue).State
            if ($EstadoPost -eq "Enabled") {
                Write-Log ".NET 3.5 instalado (offline)" "OK"
                $Script:SoftwareInstalado += ".NET 3.5"
                return
            } else {
                Write-Log "DISM offline no completo la instalacion, intentando via Windows Update..." "WARN"
            }
        } catch {
            Write-Log "Error DISM offline: $($_.Exception.Message). Intentando via Windows Update..." "WARN"
        }
    } else {
        Write-Log "Carpeta offline no encontrada ($RutaDotNet). Intentando via Windows Update..." "WARN"
    }

    # Fallback: instalar via Windows Update (requiere internet)
    Write-Log "Instalando .NET 3.5 via Windows Update..."
    Enable-WindowsOptionalFeature -Online -FeatureName "NetFx3" -All -NoRestart -ErrorAction SilentlyContinue | Out-Null

    $EstadoFinal = (Get-WindowsOptionalFeature -Online -FeatureName "NetFx3" -ErrorAction SilentlyContinue).State
    if ($EstadoFinal -eq "Enabled") {
        Write-Log ".NET 3.5 instalado (online)" "OK"
    } else {
        Write-Log "No se pudo instalar .NET 3.5" "ERROR"
    }
    $Script:SoftwareInstalado += ".NET 3.5"
}

function Install-AcrobatReader {
    Write-StepHeader -Step 12 -Title "INSTALANDO ACROBAT READER"
    Show-ProgressCosmos -Step 12

    # Verificar si ya esta instalado
    $AcrobatPaths = @(
        "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    )
    foreach ($path in $AcrobatPaths) {
        if (Test-Path $path) {
            Write-Log "Acrobat Reader ya esta instalado" "INFO"
            $Script:SoftwareInstalado += "Acrobat"
            return
        }
    }

    # Buscar instalador
    $Instalador = "$RutaAcrobat\Reader_es_install.exe"
    if (-not (Test-Path $Instalador)) {
        $Instalador = Find-Installer -Carpeta $RutaAcrobat -Filtro "*.exe"
    }

    if ($Instalador -and (Test-Path $Instalador)) {
        Write-Log "Instalando Acrobat Reader (sin McAfee)..."
        Start-Process -FilePath $Instalador -ArgumentList "/sAll /rs /msi EULA_ACCEPT=YES DISABLE_OPTIONAL_OFFER=YES SUPPRESS_APP_LAUNCH=YES" -Wait -NoNewWindow

        # Verificar instalacion
        $Instalado = $false
        foreach ($path in $AcrobatPaths) {
            if (Test-Path $path) { $Instalado = $true; break }
        }

        if ($Instalado) {
            Write-Log "Acrobat Reader instalado correctamente" "OK"
        } else {
            Write-Log "Acrobat Reader puede estar instalandose en segundo plano" "WARN"
        }

        # Desinstalar McAfee si se colo
        $McAfee = Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*McAfee*" }
        if ($McAfee) {
            Write-Log "Detectado McAfee, desinstalando..." "INFO"
            $McAfee | ForEach-Object { $_.Uninstall() | Out-Null }
            Write-Log "McAfee eliminado" "OK"
        }

        $Script:SoftwareInstalado += "Acrobat"
    } else {
        Write-Log "No se encontro instalador de Acrobat en: $RutaAcrobat" "ERROR"
    }
}

function Install-Chrome {
    Write-StepHeader -Step 13 -Title "INSTALANDO GOOGLE CHROME"
    Show-ProgressCosmos -Step 13

    $ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"

    if (-not (Test-Path $ChromePath)) {
        # Buscar instalador
        $Instalador = "$RutaChrome\ChromeSetup.exe"
        if (-not (Test-Path $Instalador)) {
            $Instalador = Find-Installer -Carpeta $RutaChrome -Filtro "*.exe"
        }

        if ($Instalador -and (Test-Path $Instalador)) {
            Write-Log "Instalando Chrome..."
            Start-Process -FilePath $Instalador -ArgumentList "/silent /install" -Wait -NoNewWindow
            Start-Sleep -Seconds 5
            Write-Log "Chrome instalado" "OK"
        } else {
            Write-Log "No se encontro instalador de Chrome" "ERROR"
            return
        }
    } else {
        Write-Log "Chrome ya esta instalado" "INFO"
    }

    # Crear acceso directo en escritorio
    $Desktop = "C:\Users\Public\Desktop"
    $ShortcutPath = "$Desktop\Google Chrome.lnk"
    if (-not (Test-Path $ShortcutPath)) {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
        $Shortcut.TargetPath = $ChromePath
        $Shortcut.WorkingDirectory = "C:\Program Files\Google\Chrome\Application"
        $Shortcut.IconLocation = "$ChromePath,0"
        $Shortcut.Save()
        Write-Log "Acceso directo de Chrome en escritorio" "OK"
    }

    # Anclar Chrome a la barra de tareas
    Pin-ToTaskbar -ShortcutPath $ShortcutPath -AppName "Google Chrome"

    # Establecer Chrome como navegador predeterminado
    Write-Log "Configurando Chrome como navegador predeterminado..."

    # Metodo 1: Usar el argumento de Chrome
    Start-Process -FilePath "$ChromePath" -ArgumentList "--make-default-browser" -NoNewWindow -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    # Metodo 2: Configurar asociaciones en registro
    $ChromeProgId = "ChromeHTML"
    $Extensions = @(".htm", ".html", ".shtml", ".xht", ".xhtml")
    $Protocols = @("http", "https")

    foreach ($ext in $Extensions) {
        $RegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext\UserChoice"
        try {
            if (Test-Path $RegPath) { Remove-Item -Path $RegPath -Force -ErrorAction SilentlyContinue }
        } catch {}
    }

    # Configurar como navegador predeterminado via Settings
    Start-Process "ms-settings:defaultapps" -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Stop-Process -Name "SystemSettings" -Force -ErrorAction SilentlyContinue

    Write-Log "Chrome configurado como navegador predeterminado" "OK"
    $Script:SoftwareInstalado += "Chrome (default)"
}

function Install-Office {
    Write-StepHeader -Step 15 -Title "INSTALANDO OFFICE 2007 (Word, Excel, PowerPoint)"
    Show-ProgressCosmos -Step 15

    # Ruta exacta del setup de Office 2007
    $SetupPath = "$RutaOffice\Ofice2007\setup.exe"
    $SerialFile = "$RutaOffice\Ofice2007\SERIAL.txt"

    # Leer serial del archivo si existe (formato: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)
    $SerialOffice = ""
    if (Test-Path $SerialFile) {
        $ContenidoSerial = Get-Content $SerialFile -First 1
        # Extraer solo el serial (quitar espacios y numeros extra al final)
        $SerialOffice = ($ContenidoSerial -split '\s+')[0].Trim() -replace '-', ''
        Write-Log "Serial de Office: $($SerialOffice.Substring(0,5))..."
    }

    if (Test-Path $SetupPath) {
        # Crear archivo de configuracion para instalar solo Word, Excel, PowerPoint
        $ConfigXML = @"
<Configuration Product="Enterprise">
    <Display Level="basic" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" />
    <PIDKEY Value="$SerialOffice" />
    <OptionState Id="WORDFiles" State="local" Children="force" />
    <OptionState Id="EXCELFiles" State="local" Children="force" />
    <OptionState Id="PPTFiles" State="local" Children="force" />
    <OptionState Id="OUTLOOKFiles" State="absent" Children="force" />
    <OptionState Id="ACCESSFiles" State="absent" Children="force" />
    <OptionState Id="PUBFiles" State="absent" Children="force" />
    <OptionState Id="ONOTEFiles" State="absent" Children="force" />
    <OptionState Id="GROOVEFiles" State="absent" Children="force" />
    <OptionState Id="INFOPATHFiles" State="absent" Children="force" />
</Configuration>
"@
        $ConfigPath = "$env:TEMP\Office2007Config.xml"
        $ConfigXML | Out-File -FilePath $ConfigPath -Encoding UTF8

        Write-Log "Instalando Office 2007 (Word, Excel, PowerPoint)..."
        Start-Process -FilePath $SetupPath -ArgumentList "/config `"$ConfigPath`"" -Wait -NoNewWindow

        Write-Log "Office 2007 instalado" "OK"
        $Script:SoftwareInstalado += "Office 2007"
    } else {
        Write-Log "No se encontro setup.exe de Office en: $SetupPath" "ERROR"
    }
}

function Install-Dedalus {
    Write-StepHeader -Step 16 -Title "INSTALANDO DEDALUS EXPEDIENTE CLINICO"
    Show-ProgressCosmos -Step 16

    $Netlogon = "$RutaDedalus\netlogon6.bat"

    if (Test-Path $Netlogon) {
        # Crear carpeta Dedalus si no existe
        if (-not (Test-Path "C:\Dedalus")) {
            New-Item -ItemType Directory -Path "C:\Dedalus" -Force | Out-Null
        }

        Write-Log "Ejecutando netlogon6.bat..."
        Start-Process cmd -ArgumentList "/c `"$Netlogon`"" -Wait -NoNewWindow
        Write-Log "Dedalus instalado" "OK"
        $Script:SoftwareInstalado += "Dedalus"
    } else {
        Write-Log "No se encontro netlogon6.bat en: $RutaDedalus" "ERROR"
    }
}

function Add-DedalusSyncStartup {
    Write-StepHeader -Step 17 -Title "AGREGANDO SINCRONIZADOR AL INICIO Y ACCESOS"
    Show-ProgressCosmos -Step 17

    $SyncBat = "$RutaDedalus\sync_xhis6_startup.bat"

    if (Test-Path $SyncBat) {
        # Copiar el archivo al equipo local
        $LocalSync = "C:\Dedalus\sync_xhis6_startup.bat"
        Copy-Item -Path $SyncBat -Destination $LocalSync -Force

        # Crear acceso directo en la carpeta de Inicio para todos los usuarios
        $StartupFolder = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
        $ShortcutPath = "$StartupFolder\Dedalus Sync.lnk"

        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
        $Shortcut.TargetPath = $LocalSync
        $Shortcut.WorkingDirectory = "C:\Dedalus"
        $Shortcut.Description = "Sincronizador Dedalus Expediente Clinico"
        $Shortcut.Save()

        Write-Log "Sincronizador agregado al inicio de Windows" "OK"
        $Script:SoftwareInstalado += "Sync Dedalus"
    } else {
        Write-Log "No se encontro sync_xhis6_startup.bat en: $RutaDedalus" "WARN"
    }

    # Anclar accesos de Dedalus/xHIS a la barra de tareas
    Add-DedalusToTaskbar
}

function Add-DedalusToTaskbar {
    Write-Host "  Anclando accesos de Expediente Clinico a barra de tareas..." -ForegroundColor Cyan

    $xHISFolder = "C:\Dedalus\xHIS"
    $Desktop = "C:\Users\Public\Desktop"
    $TaskBarFolder = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

    if (-not (Test-Path $TaskBarFolder)) {
        New-Item -ItemType Directory -Path $TaskBarFolder -Force | Out-Null
    }

    # Buscar el acceso directo xHIS v6
    $xHISShortcut = "$xHISFolder\xHIS v6.lnk"
    if (Test-Path $xHISShortcut) {
        # Copiar a escritorio
        Copy-Item -Path $xHISShortcut -Destination $Desktop -Force -ErrorAction SilentlyContinue
        # Copiar a barra de tareas
        Copy-Item -Path $xHISShortcut -Destination $TaskBarFolder -Force -ErrorAction SilentlyContinue
        Write-Log "xHIS v6 agregado a escritorio y barra de tareas" "OK"
    } else {
        # Crear acceso directo para appl_generic si no existe el .lnk
        $ApplGeneric = "$xHISFolder\appl_generic.exe"
        if (Test-Path $ApplGeneric) {
            $ShortcutPath = "$Desktop\xHIS v6.lnk"
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
            $Shortcut.TargetPath = $ApplGeneric
            $Shortcut.WorkingDirectory = $xHISFolder
            $Shortcut.Description = "Expediente Clinico Electronico xHIS"
            $IconPath = "$xHISFolder\iconos"
            if (Test-Path "$IconPath\xhis.ico") {
                $Shortcut.IconLocation = "$IconPath\xhis.ico"
            }
            $Shortcut.Save()

            # Copiar a barra de tareas
            Copy-Item -Path $ShortcutPath -Destination $TaskBarFolder -Force -ErrorAction SilentlyContinue
            Write-Log "xHIS v6 (appl_generic) agregado a escritorio y barra de tareas" "OK"
        }
    }

    # Buscar otros accesos de Dedalus para la barra de tareas
    $OtrosAccesos = @(
        @{ Carpeta = "C:\Dedalus\EscritorioClinico"; Nombre = "Escritorio Clinico" },
        @{ Carpeta = "C:\Dedalus\xFARMA"; Nombre = "xFARMA" },
        @{ Carpeta = "C:\Dedalus\hPRESC"; Nombre = "hPRESC" }
    )

    foreach ($Acceso in $OtrosAccesos) {
        if (Test-Path $Acceso.Carpeta) {
            # Buscar acceso directo .lnk
            $LnkFile = Get-ChildItem -Path $Acceso.Carpeta -Filter "*.lnk" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($LnkFile) {
                Copy-Item -Path $LnkFile.FullName -Destination $Desktop -Force -ErrorAction SilentlyContinue
                Copy-Item -Path $LnkFile.FullName -Destination $TaskBarFolder -Force -ErrorAction SilentlyContinue
                Write-Log "$($Acceso.Nombre) agregado a barra de tareas" "OK"
            } else {
                # Buscar .exe principal
                $ExeFile = Get-ChildItem -Path $Acceso.Carpeta -Filter "appl_*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($ExeFile) {
                    $ShortcutPath = "$Desktop\$($Acceso.Nombre).lnk"
                    $WshShell = New-Object -ComObject WScript.Shell
                    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
                    $Shortcut.TargetPath = $ExeFile.FullName
                    $Shortcut.WorkingDirectory = $Acceso.Carpeta
                    $Shortcut.Save()
                    Copy-Item -Path $ShortcutPath -Destination $TaskBarFolder -Force -ErrorAction SilentlyContinue
                    Write-Log "$($Acceso.Nombre) agregado a barra de tareas" "OK"
                }
            }
        }
    }
}

function Install-Antivirus {
    Write-StepHeader -Step 18 -Title "INSTALANDO ANTIVIRUS ESET"
    Show-ProgressCosmos -Step 18

    # Verificar si ya esta instalado
    $ESETService = Get-Service -Name "ekrn" -ErrorAction SilentlyContinue
    if ($ESETService) {
        Write-Log "ESET ya esta instalado" "INFO"
        $Script:SoftwareInstalado += "ESET"
        return
    }

    # Buscar instalador con nombre exacto
    $Instalador = "$RutaAntivirus\PROTECT_Installer_x64_es_CL 2.exe"
    if (-not (Test-Path $Instalador)) {
        $Instalador = Find-Installer -Carpeta $RutaAntivirus -Filtro "*.exe"
    }

    if ($Instalador -and (Test-Path $Instalador)) {
        Write-Log "Instalando ESET PROTECT... (puede tardar varios minutos)"
        Write-Host "  Archivo: $Instalador" -ForegroundColor Gray

        $Process = Start-Process -FilePath $Instalador -ArgumentList "--silent --accepteula" -Wait -NoNewWindow -PassThru

        if ($Process.ExitCode -eq 0 -or $Process.ExitCode -eq 3010) {
            Write-Log "ESET instalado correctamente" "OK"
            $Script:SoftwareInstalado += "ESET"
        } else {
            # Verificar de todas formas
            Start-Sleep -Seconds 5
            if (Get-Service -Name "ekrn" -ErrorAction SilentlyContinue) {
                Write-Log "ESET instalado (verificado)" "OK"
                $Script:SoftwareInstalado += "ESET"
            } else {
                Write-Log "Verificar manualmente la instalacion de ESET (codigo: $($Process.ExitCode))" "WARN"
            }
        }
    } else {
        Write-Log "No se encontro instalador de ESET en: $RutaAntivirus" "ERROR"
    }
}

function Copy-AccesosDirectos {
    Write-StepHeader -Step 19 -Title "COPIANDO ACCESOS DIRECTOS"
    Show-ProgressCosmos -Step 19

    $Desktop = "C:\Users\Public\Desktop"

    if (Test-Path $RutaAccesos) {
        $Accesos = Get-ChildItem -Path $RutaAccesos -Include "*.lnk","*.url" -Recurse
        $Contador = 0
        foreach ($Acceso in $Accesos) {
            Copy-Item -Path $Acceso.FullName -Destination $Desktop -Force
            $Contador++
        }
        Write-Log "$Contador accesos directos copiados" "OK"
        $Script:SoftwareInstalado += "Accesos"
    } else {
        Write-Log "No se encontro carpeta de accesos: $RutaAccesos" "WARN"
    }
}

function Remove-AdminUsuarioActual {
    Write-StepHeader -Step 20 -Title "QUITANDO PRIVILEGIOS DE ADMINISTRADOR"
    Show-ProgressCosmos -Step 20

    $UsuarioActual = $Script:UsuarioOriginal

    # No quitar admin si es el usuario Soporte o Administrator
    if ($UsuarioActual -eq $UsuarioSoporte -or $UsuarioActual -eq "Administrator" -or $UsuarioActual -eq "Administrador") {
        Write-Log "Usuario '$UsuarioActual' mantiene privilegios de administrador" "INFO"
        return
    }

    # Verificar si el usuario actual es administrador
    $EsAdmin = Get-LocalGroupMember -Group "Administradores" -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*\$UsuarioActual" }

    if ($EsAdmin) {
        try {
            Remove-LocalGroupMember -Group "Administradores" -Member $UsuarioActual -ErrorAction Stop
            Write-Log "Privilegios de administrador removidos de '$UsuarioActual'" "OK"
            Write-Host "  IMPORTANTE: El usuario '$UsuarioActual' ya no es administrador" -ForegroundColor Yellow
            Write-Host "  Solo el usuario '$UsuarioSoporte' tiene privilegios de administrador" -ForegroundColor Yellow
        } catch {
            Write-Log "No se pudieron remover privilegios de '$UsuarioActual': $($_.Exception.Message)" "WARN"
        }
    } else {
        Write-Log "Usuario '$UsuarioActual' no era administrador" "INFO"
    }

    $Script:SoftwareInstalado += "Admin removido"
}

function Install-ReporteIP {
    Write-StepHeader -Step 21 -Title "CONFIGURANDO REPORTE AUTOMATICO DE IP"
    Show-ProgressCosmos -Step 21

    $ScriptPath = "C:\HCG_Logs\report_ip.ps1"

    # Crear el script de reporte de IP
    $ScriptContent = @'
# HCG - Reporte automatico de IP (cada 3 horas + al iniciar sesion)
$ErrorActionPreference = "SilentlyContinue"
try {
    $GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 120)
    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Ethernet*" } | Select-Object -First 1
    $WiFiAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Wi-Fi*" -or $_.Name -like "*WiFi*" -or $_.Name -like "*Wireless*" } | Select-Object -First 1
    $MACEthernet = ""
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace "-", "").ToUpper() }
    if (-not $MACEthernet) { exit }
    $IPEthernet = ""
    if ($EthAdapter -and $EthAdapter.Status -eq "Up") {
        $IPEthernet = (Get-NetIPAddress -InterfaceIndex $EthAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1).IPAddress
    }
    $IPWiFi = ""
    if ($WiFiAdapter -and $WiFiAdapter.Status -eq "Up") {
        $IPWiFi = (Get-NetIPAddress -InterfaceIndex $WiFiAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1).IPAddress
    }
    $SSIDWiFi = ""
    if ($WiFiAdapter -and $WiFiAdapter.Status -eq "Up") {
        $NetshOutput = netsh wlan show interfaces 2>$null
        if ($NetshOutput) {
            $SSIDLine = $NetshOutput | Select-String "^\s+SSID\s+:" | Select-Object -First 1
            if ($SSIDLine) { $SSIDWiFi = ($SSIDLine.ToString() -replace "^\s+SSID\s+:\s+", "").Trim() }
        }
    }
    if (-not $IPEthernet -and -not $IPWiFi) { exit }
    $TestOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) { $TestOK = Test-Connection -ComputerName "dns.google" -Count 1 -Quiet -ErrorAction SilentlyContinue }
    if (-not $TestOK) { exit }
    $Body = @{
        Accion = "ip"; MACEthernet = $MACEthernet; IPEthernet = $IPEthernet
        IPWiFi = $IPWiFi; SSIDWiFi = $SSIDWiFi; NombreEquipo = $env:COMPUTERNAME
        FechaReporte = (Get-Date -Format "dd/MM/yyyy HH:mm")
    } | ConvertTo-Json
    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 30 | Out-Null
            break
        } catch { if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 10 -Maximum 30) } }
    }
} catch { exit }
'@

    if (-not (Test-Path "C:\HCG_Logs")) { New-Item -ItemType Directory -Path "C:\HCG_Logs" -Force | Out-Null }
    $ScriptContent | Out-File -FilePath $ScriptPath -Encoding UTF8 -Force
    Write-Log "Script de reporte de IP creado en $ScriptPath" "OK"

    # Crear tarea programada: cada 3 horas + al iniciar sesion
    $TaskName = "HCG_ReporteIP"

    # Eliminar tarea si ya existe
    schtasks /delete /tn $TaskName /f 2>$null | Out-Null

    try {
        $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$ScriptPath`""

        # Trigger 1: Al iniciar sesion (cualquier usuario), delay 60 segundos
        $TriggerLogon = New-ScheduledTaskTrigger -AtLogOn
        $TriggerLogon.Delay = "PT60S"

        # Trigger 2: Cada 3 horas, indefinidamente
        $TriggerRepeat = New-ScheduledTaskTrigger -Once -At "00:00" -RepetitionInterval (New-TimeSpan -Hours 3)

        $Settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
            -MultipleInstances IgnoreNew

        Register-ScheduledTask -TaskName $TaskName -Action $Action `
            -Trigger @($TriggerLogon, $TriggerRepeat) `
            -Settings $Settings `
            -Description "HCG - Reporte automatico de IP cada 3 horas" `
            -RunLevel Limited -Force | Out-Null

        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Write-Log "Tarea programada '$TaskName' creada (cada 3 horas + al iniciar sesion)" "OK"
        } else {
            Write-Log "No se pudo crear la tarea programada" "WARN"
        }
    } catch {
        Write-Log "Error al crear tarea programada: $($_.Exception.Message)" "WARN"
    }

    $Script:SoftwareInstalado += "Reporte IP"
}

function Install-ReporteSistema {
    Write-StepHeader -Step 22 -Title "CONFIGURANDO REPORTE DE SISTEMA"
    Show-ProgressCosmos -Step 22

    $ScriptPath = "C:\HCG_Logs\report_system.ps1"

    # Crear el script de reporte de sistema
    $ScriptContent = @'
# HCG - Reporte de sistema y limpieza automatica (cada inicio de sesion)
$ErrorActionPreference = "SilentlyContinue"
try {
    $GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 180)

    # --- Identificar equipo por MAC Ethernet ---
    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*Ethernet*" } | Select-Object -First 1
    $MACEthernet = ""
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace "-", "").ToUpper() }
    if (-not $MACEthernet) { exit }

    # === 1. LIMPIEZA DE ARCHIVOS TEMPORALES (solo archivos > 1 dia) ===
    $BytesLimpiados = 0
    $TempFolders = @("$env:TEMP", "C:\Windows\Temp", "$env:LOCALAPPDATA\Temp")
    foreach ($folder in $TempFolders) {
        if (Test-Path $folder) {
            Get-ChildItem -Path $folder -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
                ForEach-Object {
                    $BytesLimpiados += $_.Length
                    Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
                }
        }
    }
    if (Test-Path "C:\Windows\Prefetch") {
        Get-ChildItem "C:\Windows\Prefetch" -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } |
            ForEach-Object {
                $BytesLimpiados += $_.Length
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            }
    }
    $WUFolder = "C:\Windows\SoftwareDistribution\Download"
    if (Test-Path $WUFolder) {
        Get-ChildItem $WUFolder -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
            ForEach-Object {
                $BytesLimpiados += $_.Length
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            }
    }
    $MBLimpiados = [math]::Round($BytesLimpiados / 1MB, 1)

    # === 2. IMPRESORAS INSTALADAS ===
    $PrinterList = @()
    Get-Printer -ErrorAction SilentlyContinue | ForEach-Object {
        $Name = $_.Name; $PortName = $_.PortName; $Type = "Local"; $IP = ""
        if ($PortName -like "*USB*") { $Type = "USB" }
        elseif ($PortName -match "\d+\.\d+\.\d+\.\d+") {
            $Type = "Red"; $IP = [regex]::Match($PortName, "\d+\.\d+\.\d+\.\d+").Value
        } else {
            $Port = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
            if ($Port -and $Port.PrinterHostAddress) { $Type = "Red"; $IP = $Port.PrinterHostAddress }
        }
        if ($IP) { $PrinterList += "$Name [$Type - $IP]" } else { $PrinterList += "$Name [$Type]" }
    }

    # === 3. USUARIOS DEL SISTEMA ===
    $UserList = @()
    $AdminMembers = @()
    try {
        $AdminMembers = Get-LocalGroupMember -Group "Administradores" -ErrorAction SilentlyContinue |
            ForEach-Object { ($_.Name -split '\\')[-1] }
    } catch {}
    Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { $_.Enabled } | ForEach-Object {
        $UName = $_.Name; $IsAdmin = $AdminMembers -contains $UName
        if ($IsAdmin) { $UserList += "$UName [Admin]" } else { $UserList += $UName }
    }

    # === 4. APLICACIONES INSTALADAS ===
    $Apps = @()
    $RegPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
    foreach ($rp in $RegPaths) {
        Get-ItemProperty $rp -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName.Trim() -ne "" -and $_.DisplayName -notlike "Update for*" -and $_.DisplayName -notlike "Security Update*" } |
            ForEach-Object { $Apps += $_.DisplayName.Trim() }
    }
    $Apps = $Apps | Select-Object -Unique | Sort-Object

    # === 5. ACCESOS DIRECTOS DEL ESCRITORIO ===
    $Shortcuts = @()
    $DesktopPaths = @("C:\Users\Public\Desktop")
    if ($env:USERPROFILE -and (Test-Path "$env:USERPROFILE\Desktop")) { $DesktopPaths += "$env:USERPROFILE\Desktop" }
    foreach ($dp in $DesktopPaths) {
        if (Test-Path $dp) {
            Get-ChildItem $dp -Filter "*.lnk" -ErrorAction SilentlyContinue | ForEach-Object { $Shortcuts += $_.BaseName }
            Get-ChildItem $dp -Filter "*.url" -ErrorAction SilentlyContinue | ForEach-Object { $Shortcuts += $_.BaseName + " (web)" }
        }
    }
    $Shortcuts = $Shortcuts | Select-Object -Unique | Sort-Object

    # === 6. ESPACIO LIBRE EN DISCO ===
    $Disco = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    $EspacioLibreGB = if ($Disco) { [math]::Round($Disco.FreeSpace / 1GB, 1) } else { 0 }
    $EspacioTotalGB = if ($Disco) { [math]::Round($Disco.Size / 1GB, 0) } else { 0 }

    # === 7. VERIFICAR INTERNET Y ENVIAR (con reintentos) ===
    $TestOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) { $TestOK = Test-Connection -ComputerName "dns.google" -Count 1 -Quiet -ErrorAction SilentlyContinue }
    if (-not $TestOK) { exit }

    $Body = @{
        Accion            = "sistema"
        MACEthernet       = $MACEthernet
        NombreEquipo      = $env:COMPUTERNAME
        Impresoras        = ($PrinterList -join " | ")
        Usuarios          = ($UserList -join " | ")
        AppsInstaladas    = ($Apps -join " | ")
        AccesosEscritorio = ($Shortcuts -join " | ")
        EspacioLibreGB    = "$EspacioLibreGB / $EspacioTotalGB GB"
        MBLimpiados       = $MBLimpiados
        FechaReporte      = (Get-Date -Format "dd/MM/yyyy HH:mm")
    } | ConvertTo-Json -Depth 3

    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 60 | Out-Null
            break
        } catch { if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 15 -Maximum 45) } }
    }
} catch { exit }
'@

    if (-not (Test-Path "C:\HCG_Logs")) { New-Item -ItemType Directory -Path "C:\HCG_Logs" -Force | Out-Null }
    $ScriptContent | Out-File -FilePath $ScriptPath -Encoding UTF8 -Force
    Write-Log "Script de reporte de sistema creado en $ScriptPath" "OK"

    # Crear tarea programada: solo al iniciar sesion, delay 2 minutos
    $TaskName = "HCG_ReporteSistema"

    # Eliminar tarea si ya existe
    schtasks /delete /tn $TaskName /f 2>$null | Out-Null

    try {
        $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$ScriptPath`""

        # Trigger: Al iniciar sesion (cualquier usuario), delay 120 segundos
        $TriggerLogon = New-ScheduledTaskTrigger -AtLogOn
        $TriggerLogon.Delay = "PT120S"

        $Settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
            -MultipleInstances IgnoreNew

        Register-ScheduledTask -TaskName $TaskName -Action $Action `
            -Trigger $TriggerLogon `
            -Settings $Settings `
            -Description "HCG - Reporte de sistema y limpieza al iniciar sesion" `
            -RunLevel Limited -Force | Out-Null

        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Write-Log "Tarea programada '$TaskName' creada (al iniciar sesion, delay 2 min)" "OK"
        } else {
            Write-Log "No se pudo crear la tarea programada de sistema" "WARN"
        }
    } catch {
        Write-Log "Error al crear tarea programada: $($_.Exception.Message)" "WARN"
    }

    $Script:SoftwareInstalado += "Reporte Sistema"
}

function Install-ReporteDiagnostico {
    Write-StepHeader -Step 23 -Title "CONFIGURANDO REPORTE DE DIAGNOSTICO DE SALUD"
    Show-ProgressCosmos -Step 23

    $ScriptPath = "C:\HCG_Logs\report_diagnostico.ps1"

    # Crear el script de diagnostico de salud
    $ScriptContent = @'
# HCG - Reporte de diagnostico de salud (cada 4 horas + al iniciar sesion)
$ErrorActionPreference = "SilentlyContinue"
try {
    $GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 120)

    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*Ethernet*" } | Select-Object -First 1
    $MACEthernet = ""
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace "-", "").ToUpper() }
    if (-not $MACEthernet) { exit }

    $TestOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) { $TestOK = Test-Connection -ComputerName "dns.google" -Count 1 -Quiet -ErrorAction SilentlyContinue }
    if (-not $TestOK) { exit }

    # --- RAM ---
    $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $RAMTotalGB = [math]::Round($OS.TotalVisibleMemorySize / 1MB, 1)
    $RAMLibreGB = [math]::Round($OS.FreePhysicalMemory / 1MB, 1)
    $RAMUsadaGB = [math]::Round($RAMTotalGB - $RAMLibreGB, 1)
    $RAMPct = [math]::Round(($RAMUsadaGB / $RAMTotalGB) * 100, 0)

    # --- Top 5 procesos ---
    $Top5 = Get-Process -ErrorAction SilentlyContinue |
        Sort-Object WorkingSet64 -Descending | Select-Object -First 5 |
        ForEach-Object { "$($_.ProcessName) ($([math]::Round($_.WorkingSet64 / 1MB, 0)) MB)" }
    $Top5Str = $Top5 -join " | "

    # --- Chrome ---
    $ChromeProcs = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    $ChromeMB = 0; $ChromeCount = 0
    if ($ChromeProcs) {
        $ChromeCount = @($ChromeProcs).Count
        $ChromeMB = [math]::Round(($ChromeProcs | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    }

    # --- Dedalus ---
    $DedalusProcs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like "*dedalus*" }
    $DedalusMB = 0; $DedalusCount = 0
    if ($DedalusProcs) {
        $DedalusCount = @($DedalusProcs).Count
        $DedalusMB = [math]::Round(($DedalusProcs | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    }

    $TotalProcs = @(Get-Process -ErrorAction SilentlyContinue).Count
    $CPU = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $CPUPct = if ($CPU) { $CPU.LoadPercentage } else { 0 }
    if (-not $CPUPct) { $CPUPct = 0 }

    $PageFile = Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue | Select-Object -First 1
    $PageFileUsadoMB = if ($PageFile) { $PageFile.CurrentUsage } else { 0 }
    $PageFileTotalMB = if ($PageFile) { $PageFile.AllocatedBaseSize } else { 0 }

    $LastBoot = $OS.LastBootUpTime
    $UptimeDias = if ($LastBoot) { [math]::Round(((Get-Date) - $LastBoot).TotalDays, 1) } else { 0 }

    $Disco = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    $DiscoLibreGB = if ($Disco) { [math]::Round($Disco.FreeSpace / 1GB, 1) } else { 0 }
    $DiscoTotalGB = if ($Disco) { [math]::Round($Disco.Size / 1GB, 0) } else { 0 }

    # --- Estado ---
    $Estado = "OK"; $Recomendaciones = @()
    if ($RAMPct -gt 85) {
        $Estado = "Critico"; $Recomendaciones += "RAM critica ($RAMPct%). Se recomienda ampliar memoria"
    } elseif ($RAMPct -gt 70) {
        $Estado = "Atencion"; $Recomendaciones += "RAM elevada ($RAMPct%). Monitorear uso. Considerar ampliacion"
    } else {
        $Recomendaciones += "Equipo operando con recursos suficientes"
    }
    if ($ChromeMB -gt 1500) { $Recomendaciones += "Chrome consumiendo $ChromeMB MB. Reducir pestanas" }
    if ($UptimeDias -gt 15) { $Recomendaciones += "Sin reinicio hace $UptimeDias dias. Reiniciar pronto" }
    if ($DiscoLibreGB -lt 20) { $Recomendaciones += "Disco bajo: $DiscoLibreGB GB libres. Liberar espacio" }
    $RecomendacionStr = $Recomendaciones -join " | "

    $Body = @{
        Accion = "diagnostico"; MACEthernet = $MACEthernet; NombreEquipo = $env:COMPUTERNAME
        RAMTotalGB = $RAMTotalGB; RAMUsadaGB = $RAMUsadaGB; RAMLibreGB = $RAMLibreGB; RAMPct = $RAMPct
        Top5Procesos = $Top5Str; ChromeMB = $ChromeMB; ChromeProcs = $ChromeCount
        DedalusMB = $DedalusMB; DedalusProcs = $DedalusCount; TotalProcs = $TotalProcs
        CPUPct = $CPUPct; PageFileUsado = $PageFileUsadoMB; PageFileTotal = $PageFileTotalMB
        UptimeDias = $UptimeDias; DiscoLibreGB = "$DiscoLibreGB / $DiscoTotalGB GB"
        Estado = $Estado; Recomendacion = $RecomendacionStr
        FechaReporte = (Get-Date -Format "dd/MM/yyyy HH:mm")
    } | ConvertTo-Json
    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 30 | Out-Null
            break
        } catch { if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 10 -Maximum 30) } }
    }
} catch { exit }
'@

    if (-not (Test-Path "C:\HCG_Logs")) { New-Item -ItemType Directory -Path "C:\HCG_Logs" -Force | Out-Null }
    $ScriptContent | Out-File -FilePath $ScriptPath -Encoding UTF8 -Force
    Write-Log "Script de diagnostico creado en $ScriptPath" "OK"

    # Crear tarea programada: cada 4 horas + al iniciar sesion (delay 3 min)
    $TaskName = "HCG_ReporteDiagnostico"

    # Eliminar tarea si ya existe
    schtasks /delete /tn $TaskName /f 2>$null | Out-Null

    try {
        $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$ScriptPath`""

        # Trigger 1: Al iniciar sesion (cualquier usuario), delay 180 segundos
        $TriggerLogon = New-ScheduledTaskTrigger -AtLogOn
        $TriggerLogon.Delay = "PT180S"

        # Trigger 2: Cada 4 horas, indefinidamente
        $TriggerRepeat = New-ScheduledTaskTrigger -Once -At "00:00" -RepetitionInterval (New-TimeSpan -Hours 4)

        $Settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
            -MultipleInstances IgnoreNew

        Register-ScheduledTask -TaskName $TaskName -Action $Action `
            -Trigger @($TriggerLogon, $TriggerRepeat) `
            -Settings $Settings `
            -Description "HCG - Diagnostico de salud cada 4 horas" `
            -RunLevel Limited -Force | Out-Null

        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Write-Log "Tarea programada '$TaskName' creada (cada 4 horas + al iniciar sesion, delay 3 min)" "OK"
        } else {
            Write-Log "No se pudo crear la tarea programada de diagnostico" "WARN"
        }
    } catch {
        Write-Log "Error al crear tarea programada: $($_.Exception.Message)" "WARN"
    }

    $Script:SoftwareInstalado += "Reporte Diagnostico"
}

# =============================================================================
# EJECUCION PRINCIPAL
# =============================================================================

Show-Banner

Write-Host "  $([char]0x2605) Ingresa el numero de inventario ST (5 digitos)" -ForegroundColor DarkYellow
$NumInventario = Read-Host "  Numero"

if ($NumInventario -notmatch '^\d{5}$') {
    Write-Host "`n  $([char]0x2716) [ERROR] Debe ser de 5 digitos" -ForegroundColor Red
    Read-Host "  Presiona Enter para salir"
    exit
}

Show-CosmosAnimation -Message "Preparando configuracion cosmica..."
Write-Host ""
Write-Host "  $([char]0x2734) Iniciando configuracion para equipo: $NumInventario" -ForegroundColor Cyan
Write-Separator

# Obtener datos del equipo
$Datos = Get-DatosEquipo

# PASO 1: Registrar inicio
Send-DatosInicio -InvST $NumInventario -Datos $Datos
Play-StepSound

# PASO 2: Conectar al servidor
if (-not (Connect-Servidor)) {
    Write-Host "`n  $([char]0x2716) [ERROR] No se pudo conectar al servidor. Abortando." -ForegroundColor Red
    Read-Host "  Presiona Enter para salir"
    exit
}
Play-StepSound

# PASOS 3-22: Instalacion y configuracion
Remove-OfficePrevio; Play-StepSound
New-UsuarioSoporte; Play-StepSound
New-UsuarioEquipo -NumInventario $NumInventario; Play-StepSound
Set-ImagenesUsuarios -NumInventario $NumInventario; Play-StepSound
Set-RedPrivada; Play-StepSound
Set-HoraAutomatica; Play-StepSound
Set-TemaOscuro; Play-StepSound
Install-WinRAR; Play-StepSound
Install-DotNet35; Play-StepSound
Install-AcrobatReader; Play-StepSound
Install-Chrome; Play-StepSound
Set-FondoPantalla; Play-StepSound
Install-Office; Play-StepSound
Install-Dedalus; Play-StepSound
Add-DedalusSyncStartup; Play-StepSound
Install-Antivirus; Play-StepSound
Copy-AccesosDirectos; Play-StepSound
Remove-AdminUsuarioActual; Play-StepSound
Install-ReporteIP; Play-StepSound
Install-ReporteSistema; Play-StepSound
Install-ReporteDiagnostico; Play-StepSound

# Renombrar equipo
$NuevoNombre = "PC-$NumInventario"
Rename-Computer -NewName $NuevoNombre -Force -ErrorAction SilentlyContinue
Write-Log "Equipo renombrado a: $NuevoNombre" "OK"

# PASO 23: Actualizar Google Sheets
Send-DatosFin -InvST $NumInventario
Play-StepSound

# PASO 24: Registrar inventario de software
Send-SoftwareInfo -InvST $NumInventario
Play-StepSound

# Mostrar resumen final cosmico
$Star = [char]0x2605
$Spark = [char]0x2734
$Arrow = [char]0x2192

# Animacion de celebracion
Show-CosmosAnimation -Message "Victoria cosmica alcanzada!"

# Melodia de victoria
Play-VictorySound

Show-ProgressCosmos -Step 25 -Total 25

Write-Host ""
Write-Host "  $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star" -ForegroundColor DarkYellow
Write-Host ""
Write-Host "        $Spark  CONFIGURACION COSMICA COMPLETADA  $Spark" -ForegroundColor Green
Write-Host "        $Spark  EL COSMO HA SIDO ENCENDIDO!       $Spark" -ForegroundColor Cyan
Write-Host ""
Write-Host "  $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star" -ForegroundColor DarkYellow
Write-Host ""
Write-Host "  $Spark DATOS DEL EQUIPO:" -ForegroundColor DarkYellow
Write-Host "  $Arrow Nombre:        $NuevoNombre" -ForegroundColor White
Write-Host "  $Arrow No. Serie:     $($Datos.Serie)" -ForegroundColor White
Write-Host "  $Arrow MAC Ethernet:  $($Datos.MACEthernet)" -ForegroundColor White
Write-Host "  $Arrow MAC WiFi:      $($Datos.MACWiFi)" -ForegroundColor White
Write-Host ""
Write-Host "  $Spark SOFTWARE INSTALADO:" -ForegroundColor DarkYellow
Write-Host "  $Arrow $($Script:SoftwareInstalado -join ', ')" -ForegroundColor Cyan
Write-Host ""
Write-Host "  $Spark SEGURIDAD:" -ForegroundColor DarkYellow
Write-Host "  $Arrow Usuario normal: $NumInventario (auto-login, estandar)" -ForegroundColor White
Write-Host "  $Arrow Usuario admin: $UsuarioSoporte (solo soporte tecnico)" -ForegroundColor White
Write-Host "  $Arrow Usuario '$($Script:UsuarioOriginal)' sin privilegios de admin" -ForegroundColor White
Write-Host ""
Write-Host "  $Spark REPORTES AUTOMATICOS:" -ForegroundColor DarkYellow
Write-Host "  $Arrow IP: cada 3 horas + al iniciar sesion" -ForegroundColor White
Write-Host "  $Arrow Sistema: impresoras, apps, usuarios, limpieza (al iniciar sesion)" -ForegroundColor White
Write-Host "  $Arrow Diagnostico: RAM, CPU, procesos, disco (cada 4 horas)" -ForegroundColor White
Write-Host ""
Write-Host "  $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star" -ForegroundColor DarkYellow
Write-Host ""
Write-Host "  $Spark IMPORTANTE: Reinicia el equipo para aplicar todos los cambios" -ForegroundColor Yellow
Write-Host ""
Write-Host "  $Star  Los Caballeros de Informatica protegen este equipo  $Star" -ForegroundColor Magenta
Write-Host "  $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star" -ForegroundColor DarkYellow
Write-Host ""

$R = Read-Host "  $Star Reiniciar ahora? (S/N)"
if ($R -eq "S" -or $R -eq "s") { Restart-Computer -Force }
