# =============================================================================
# HCG - INSTALAR TAREAS DE REPORTE (IP + Sistema)
# =============================================================================
# Ejecutar en equipos que ya fueron configurados pero no tienen los reportes
# Crea los scripts y las tareas programadas de reporte automatico
# Se eleva automaticamente como Administrador
# =============================================================================

# --- Auto-elevacion como Administrador ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  HCG - INSTALADOR DE REPORTES AUTOMATICOS" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

# Crear carpeta si no existe
if (-not (Test-Path "C:\HCG_Logs")) {
    New-Item -ItemType Directory -Path "C:\HCG_Logs" -Force | Out-Null
}

# =============================================================================
# 1. SCRIPT DE REPORTE DE IP
# =============================================================================

Write-Host "  [1/3] Creando reporte de IP..." -ForegroundColor Yellow

$IPScript = @'
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
    if (-not $TestOK) { try { $null = Invoke-WebRequest -Uri "https://www.google.com" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop; $TestOK = $true } catch {} }
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

$IPScript | Out-File -FilePath "C:\HCG_Logs\report_ip.ps1" -Encoding UTF8 -Force

# Crear tarea programada de IP (Register -Force reemplaza si ya existe)
try {
    $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"C:\HCG_Logs\report_ip.ps1`""

    $TriggerLogon = New-ScheduledTaskTrigger -AtLogOn
    $TriggerLogon.Delay = "PT60S"

    $TriggerRepeat = New-ScheduledTaskTrigger -Once -At "00:00" -RepetitionInterval (New-TimeSpan -Hours 3)

    $Settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
        -MultipleInstances IgnoreNew

    Register-ScheduledTask -TaskName "HCG_ReporteIP" -Action $Action `
        -Trigger @($TriggerLogon, $TriggerRepeat) `
        -Settings $Settings `
        -Description "HCG - Reporte automatico de IP cada 3 horas" `
        -RunLevel Limited -Force | Out-Null

    if (Get-ScheduledTask -TaskName "HCG_ReporteIP" -ErrorAction SilentlyContinue) {
        Write-Host "  [OK] Tarea HCG_ReporteIP creada (cada 3 horas + al iniciar sesion)" -ForegroundColor Green
    } else {
        Write-Host "  [ERROR] No se pudo crear HCG_ReporteIP" -ForegroundColor Red
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
}

# =============================================================================
# 2. SCRIPT DE REPORTE DE SISTEMA
# =============================================================================

Write-Host "  [2/3] Creando reporte de sistema..." -ForegroundColor Yellow

$SysScript = @'
# HCG - Reporte de sistema y limpieza automatica (cada inicio de sesion)
$ErrorActionPreference = "SilentlyContinue"
try {
    $GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 180)

    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*Ethernet*" } | Select-Object -First 1
    $MACEthernet = ""
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace "-", "").ToUpper() }
    if (-not $MACEthernet) { exit }

    # 1. LIMPIEZA
    $BytesLimpiados = 0
    $TempFolders = @("$env:TEMP", "C:\Windows\Temp", "$env:LOCALAPPDATA\Temp")
    foreach ($folder in $TempFolders) {
        if (Test-Path $folder) {
            Get-ChildItem -Path $folder -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
                ForEach-Object {
                    $sz = $_.Length; $fp = $_.FullName
                    Remove-Item $fp -Force -ErrorAction SilentlyContinue
                    if (-not (Test-Path $fp)) { $BytesLimpiados += $sz }
                }
        }
    }
    if (Test-Path "C:\Windows\Prefetch") {
        Get-ChildItem "C:\Windows\Prefetch" -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } |
            ForEach-Object {
                $sz = $_.Length; $fp = $_.FullName
                Remove-Item $fp -Force -ErrorAction SilentlyContinue
                if (-not (Test-Path $fp)) { $BytesLimpiados += $sz }
            }
    }
    $WUFolder = "C:\Windows\SoftwareDistribution\Download"
    if (Test-Path $WUFolder) {
        Get-ChildItem $WUFolder -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
            ForEach-Object {
                $sz = $_.Length; $fp = $_.FullName
                Remove-Item $fp -Force -ErrorAction SilentlyContinue
                if (-not (Test-Path $fp)) { $BytesLimpiados += $sz }
            }
    }
    $MBLimpiados = [math]::Round($BytesLimpiados / 1MB, 1)

    # 2. IMPRESORAS
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

    # 3. USUARIOS
    $UserList = @()
    $AdminMembers = @()
    try {
        $AdminGroupName = (Get-LocalGroup -ErrorAction SilentlyContinue | Where-Object { $_.SID.Value -eq 'S-1-5-32-544' }).Name
        if ($AdminGroupName) {
            $AdminMembers = Get-LocalGroupMember -Group $AdminGroupName -ErrorAction SilentlyContinue |
                ForEach-Object { ($_.Name -split '\\')[-1] }
        }
    } catch {}
    Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { $_.Enabled } | ForEach-Object {
        $UName = $_.Name; $IsAdmin = $AdminMembers -contains $UName
        if ($IsAdmin) { $UserList += "$UName [Admin]" } else { $UserList += $UName }
    }

    # 4. APPS INSTALADAS
    $Apps = @()
    $RegPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
    foreach ($rp in $RegPaths) {
        Get-ItemProperty $rp -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName.Trim() -ne "" -and $_.DisplayName -notlike "Update for*" -and $_.DisplayName -notlike "Security Update*" } |
            ForEach-Object { $Apps += $_.DisplayName.Trim() }
    }
    $Apps = $Apps | Select-Object -Unique | Sort-Object

    # 5. ACCESOS DIRECTOS
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

    # 6. DISCO
    $Disco = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    $EspacioLibreGB = if ($Disco) { [math]::Round($Disco.FreeSpace / 1GB, 1) } else { 0 }
    $EspacioTotalGB = if ($Disco) { [math]::Round($Disco.Size / 1GB, 0) } else { 0 }

    # 7. ENVIAR
    $TestOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) { $TestOK = Test-Connection -ComputerName "dns.google" -Count 1 -Quiet -ErrorAction SilentlyContinue }
    if (-not $TestOK) { try { $null = Invoke-WebRequest -Uri "https://www.google.com" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop; $TestOK = $true } catch {} }
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

$SysScript | Out-File -FilePath "C:\HCG_Logs\report_system.ps1" -Encoding UTF8 -Force

# Crear tarea programada de Sistema (Register -Force reemplaza si ya existe)
try {
    $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"C:\HCG_Logs\report_system.ps1`""

    $TriggerLogon = New-ScheduledTaskTrigger -AtLogOn
    $TriggerLogon.Delay = "PT120S"

    $Settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
        -MultipleInstances IgnoreNew

    Register-ScheduledTask -TaskName "HCG_ReporteSistema" -Action $Action `
        -Trigger $TriggerLogon `
        -Settings $Settings `
        -Description "HCG - Reporte de sistema y limpieza al iniciar sesion" `
        -RunLevel Limited -Force | Out-Null

    if (Get-ScheduledTask -TaskName "HCG_ReporteSistema" -ErrorAction SilentlyContinue) {
        Write-Host "  [OK] Tarea HCG_ReporteSistema creada (al iniciar sesion)" -ForegroundColor Green
    } else {
        Write-Host "  [ERROR] No se pudo crear HCG_ReporteSistema" -ForegroundColor Red
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
}

# =============================================================================
# 3. SCRIPT DE REPORTE DE DIAGNOSTICO DE SALUD
# =============================================================================

Write-Host "  [3/3] Creando reporte de diagnostico de salud..." -ForegroundColor Yellow

$DiagScript = @'
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
    if (-not $TestOK) { try { $null = Invoke-WebRequest -Uri "https://www.google.com" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop; $TestOK = $true } catch {} }
    if (-not $TestOK) { exit }

    # --- RAM ---
    $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $RAMTotalGB = if ($OS.TotalVisibleMemorySize) { [math]::Round($OS.TotalVisibleMemorySize / 1MB, 1) } else { 0 }
    $RAMLibreGB = if ($OS.FreePhysicalMemory) { [math]::Round($OS.FreePhysicalMemory / 1MB, 1) } else { 0 }
    $RAMUsadaGB = [math]::Round($RAMTotalGB - $RAMLibreGB, 1)
    $RAMPct = if ($RAMTotalGB -gt 0) { [math]::Round(($RAMUsadaGB / $RAMTotalGB) * 100, 0) } else { 0 }

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

$DiagScript | Out-File -FilePath "C:\HCG_Logs\report_diagnostico.ps1" -Encoding UTF8 -Force

# Crear tarea programada de Diagnostico (Register -Force reemplaza si ya existe)
try {
    $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"C:\HCG_Logs\report_diagnostico.ps1`""

    $TriggerLogon = New-ScheduledTaskTrigger -AtLogOn
    $TriggerLogon.Delay = "PT180S"

    $TriggerRepeat = New-ScheduledTaskTrigger -Once -At "00:00" -RepetitionInterval (New-TimeSpan -Hours 4)

    $Settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
        -MultipleInstances IgnoreNew

    Register-ScheduledTask -TaskName "HCG_ReporteDiagnostico" -Action $Action `
        -Trigger @($TriggerLogon, $TriggerRepeat) `
        -Settings $Settings `
        -Description "HCG - Diagnostico de salud cada 4 horas" `
        -RunLevel Limited -Force | Out-Null

    if (Get-ScheduledTask -TaskName "HCG_ReporteDiagnostico" -ErrorAction SilentlyContinue) {
        Write-Host "  [OK] Tarea HCG_ReporteDiagnostico creada (cada 4 horas + al iniciar sesion)" -ForegroundColor Green
    } else {
        Write-Host "  [ERROR] No se pudo crear HCG_ReporteDiagnostico" -ForegroundColor Red
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
}

# =============================================================================
# RESUMEN
# =============================================================================

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  REPORTES INSTALADOS" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Scripts creados:" -ForegroundColor White
Write-Host "    C:\HCG_Logs\report_ip.ps1" -ForegroundColor Gray
Write-Host "    C:\HCG_Logs\report_system.ps1" -ForegroundColor Gray
Write-Host "    C:\HCG_Logs\report_diagnostico.ps1" -ForegroundColor Gray
Write-Host ""
Write-Host "  Tareas programadas:" -ForegroundColor White
Write-Host "    HCG_ReporteIP          - cada 3 horas + al iniciar sesion" -ForegroundColor Gray
Write-Host "    HCG_ReporteSistema     - al iniciar sesion (delay 2 min)" -ForegroundColor Gray
Write-Host "    HCG_ReporteDiagnostico - cada 4 horas + al iniciar sesion (delay 3 min)" -ForegroundColor Gray
Write-Host ""
Write-Host "  Los reportes se enviaran en el proximo inicio de sesion." -ForegroundColor Yellow
Write-Host "  Para probar ahora, reinicia el equipo." -ForegroundColor Yellow
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""

Read-Host "  Presiona Enter para salir"
