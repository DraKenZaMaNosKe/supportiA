# =============================================================================
# HCG - REPORTE DE SISTEMA Y LIMPIEZA AUTOMATICA
# =============================================================================
# Se ejecuta en cada inicio de sesion de Windows
# Limpia temporales y reporta: impresoras, usuarios, apps, accesos directos
# A prueba de errores: si falla algo, sale silenciosamente
# Con delay aleatorio y reintentos para evitar colisiones
# =============================================================================

$ErrorActionPreference = "SilentlyContinue"

try {
    $GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # --- Delay aleatorio (0-180 seg) para no saturar el servidor ---
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 180)

    # --- Identificar equipo por MAC Ethernet ---
    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*Ethernet*" } | Select-Object -First 1
    $MACEthernet = ""
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace "-", "").ToUpper() }
    if (-not $MACEthernet) { exit }

    # =========================================================================
    # 1. LIMPIEZA DE ARCHIVOS TEMPORALES (solo archivos > 1 dia)
    # =========================================================================
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

    # Limpiar Prefetch (archivos > 7 dias)
    if (Test-Path "C:\Windows\Prefetch") {
        Get-ChildItem "C:\Windows\Prefetch" -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } |
            ForEach-Object {
                $BytesLimpiados += $_.Length
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            }
    }

    # Limpiar carpeta de descargas de Windows Update
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

    # =========================================================================
    # 2. IMPRESORAS INSTALADAS
    # =========================================================================
    $PrinterList = @()
    Get-Printer -ErrorAction SilentlyContinue | ForEach-Object {
        $Name = $_.Name
        $PortName = $_.PortName
        $Type = "Local"
        $IP = ""

        if ($PortName -like "*USB*") {
            $Type = "USB"
        } elseif ($PortName -match "\d+\.\d+\.\d+\.\d+") {
            $Type = "Red"
            $IP = [regex]::Match($PortName, "\d+\.\d+\.\d+\.\d+").Value
        } else {
            $Port = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
            if ($Port -and $Port.PrinterHostAddress) {
                $Type = "Red"
                $IP = $Port.PrinterHostAddress
            }
        }

        if ($IP) {
            $PrinterList += "$Name [$Type - $IP]"
        } else {
            $PrinterList += "$Name [$Type]"
        }
    }

    # =========================================================================
    # 3. USUARIOS DEL SISTEMA
    # =========================================================================
    $UserList = @()
    $AdminMembers = @()
    try {
        $AdminMembers = Get-LocalGroupMember -Group "Administradores" -ErrorAction SilentlyContinue |
            ForEach-Object { ($_.Name -split '\\')[-1] }
    } catch {}

    Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { $_.Enabled } | ForEach-Object {
        $UName = $_.Name
        $IsAdmin = $AdminMembers -contains $UName
        if ($IsAdmin) {
            $UserList += "$UName [Admin]"
        } else {
            $UserList += $UName
        }
    }

    # =========================================================================
    # 4. APLICACIONES INSTALADAS
    # =========================================================================
    $Apps = @()
    $RegPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($rp in $RegPaths) {
        Get-ItemProperty $rp -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName.Trim() -ne "" -and $_.DisplayName -notlike "Update for*" -and $_.DisplayName -notlike "Security Update*" } |
            ForEach-Object { $Apps += $_.DisplayName.Trim() }
    }
    $Apps = $Apps | Select-Object -Unique | Sort-Object

    # =========================================================================
    # 5. ACCESOS DIRECTOS DEL ESCRITORIO
    # =========================================================================
    $Shortcuts = @()
    $DesktopPaths = @("C:\Users\Public\Desktop")

    # Agregar escritorio del usuario actual
    if ($env:USERPROFILE -and (Test-Path "$env:USERPROFILE\Desktop")) {
        $DesktopPaths += "$env:USERPROFILE\Desktop"
    }

    foreach ($dp in $DesktopPaths) {
        if (Test-Path $dp) {
            Get-ChildItem $dp -Filter "*.lnk" -ErrorAction SilentlyContinue |
                ForEach-Object { $Shortcuts += $_.BaseName }
            Get-ChildItem $dp -Filter "*.url" -ErrorAction SilentlyContinue |
                ForEach-Object { $Shortcuts += $_.BaseName + " (web)" }
        }
    }
    $Shortcuts = $Shortcuts | Select-Object -Unique | Sort-Object

    # =========================================================================
    # 6. ESPACIO LIBRE EN DISCO
    # =========================================================================
    $Disco = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    $EspacioLibreGB = if ($Disco) { [math]::Round($Disco.FreeSpace / 1GB, 1) } else { 0 }
    $EspacioTotalGB = if ($Disco) { [math]::Round($Disco.Size / 1GB, 0) } else { 0 }

    # =========================================================================
    # 7. VERIFICAR INTERNET Y ENVIAR (con reintentos)
    # =========================================================================
    $TestOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) {
        $TestOK = Test-Connection -ComputerName "dns.google" -Count 1 -Quiet -ErrorAction SilentlyContinue
    }
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

    $Enviado = $false
    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json; charset=utf-8" -TimeoutSec 60 | Out-Null
            $Enviado = $true
            break
        } catch {
            if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 15 -Maximum 45) }
        }
    }

} catch {
    # Silencio total - NUNCA interrumpir Windows
    exit
}
