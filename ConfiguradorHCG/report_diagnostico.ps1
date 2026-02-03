# =============================================================================
# HCG - REPORTE AUTOMATICO DE DIAGNOSTICO DE SALUD
# =============================================================================
# Se ejecuta cada 4 horas y al iniciar sesion de Windows
# Envia datos de RAM, CPU, procesos, disco a Google Sheets
# A prueba de errores: si no hay internet, sale silenciosamente
# Con delay aleatorio y reintentos para evitar colisiones
# =============================================================================

$ErrorActionPreference = "SilentlyContinue"

try {
    $GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # --- Delay aleatorio (0-120 seg) para no saturar el servidor ---
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 120)

    # --- Obtener MAC Ethernet para identificar equipo ---
    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*Ethernet*" } | Select-Object -First 1

    $MACEthernet = ""
    if ($EthAdapter) {
        $MACEthernet = ($EthAdapter.MacAddress -replace "-", "").ToUpper()
    }

    # Si no hay MAC Ethernet, no podemos identificar el equipo
    if (-not $MACEthernet) { exit }

    # --- Verificar conectividad a internet (ICMP + HTTP fallback) ---
    $TestOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) {
        $TestOK = Test-Connection -ComputerName "dns.google" -Count 1 -Quiet -ErrorAction SilentlyContinue
    }
    if (-not $TestOK) {
        try { $null = Invoke-WebRequest -Uri "https://www.google.com" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop; $TestOK = $true } catch {}
    }
    if (-not $TestOK) { exit }

    # =================================================================
    # RECOLECCION DE DATOS DE SALUD
    # =================================================================

    # --- RAM ---
    $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $RAMTotalGB = [math]::Round($OS.TotalVisibleMemorySize / 1MB, 1)
    $RAMLibreGB = [math]::Round($OS.FreePhysicalMemory / 1MB, 1)
    $RAMUsadaGB = [math]::Round($RAMTotalGB - $RAMLibreGB, 1)
    $RAMPct = [math]::Round(($RAMUsadaGB / $RAMTotalGB) * 100, 0)

    # --- Top 5 procesos por consumo de memoria ---
    $Top5 = Get-Process -ErrorAction SilentlyContinue |
        Sort-Object WorkingSet64 -Descending |
        Select-Object -First 5 |
        ForEach-Object {
            "$($_.ProcessName) ($([math]::Round($_.WorkingSet64 / 1MB, 0)) MB)"
        }
    $Top5Str = $Top5 -join " | "

    # --- Chrome ---
    $ChromeProcs = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    $ChromeMB = 0
    $ChromeCount = 0
    if ($ChromeProcs) {
        $ChromeCount = @($ChromeProcs).Count
        $ChromeMB = [math]::Round(($ChromeProcs | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    }

    # --- Dedalus ---
    $DedalusProcs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like "*dedalus*" }
    $DedalusMB = 0
    $DedalusCount = 0
    if ($DedalusProcs) {
        $DedalusCount = @($DedalusProcs).Count
        $DedalusMB = [math]::Round(($DedalusProcs | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    }

    # --- Total procesos ---
    $TotalProcs = @(Get-Process -ErrorAction SilentlyContinue).Count

    # --- CPU % (snapshot instantaneo) ---
    $CPU = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $CPUPct = if ($CPU) { $CPU.LoadPercentage } else { 0 }
    if (-not $CPUPct) { $CPUPct = 0 }

    # --- Page File ---
    $PageFile = Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue | Select-Object -First 1
    $PageFileUsadoMB = if ($PageFile) { $PageFile.CurrentUsage } else { 0 }
    $PageFileTotalMB = if ($PageFile) { $PageFile.AllocatedBaseSize } else { 0 }

    # --- Uptime ---
    $LastBoot = $OS.LastBootUpTime
    $UptimeDias = if ($LastBoot) { [math]::Round(((Get-Date) - $LastBoot).TotalDays, 1) } else { 0 }

    # --- Disco C: ---
    $Disco = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    $DiscoLibreGB = if ($Disco) { [math]::Round($Disco.FreeSpace / 1GB, 1) } else { 0 }
    $DiscoTotalGB = if ($Disco) { [math]::Round($Disco.Size / 1GB, 0) } else { 0 }

    # =================================================================
    # CLASIFICACION DE ESTADO Y RECOMENDACIONES
    # =================================================================

    $Estado = "OK"
    $Recomendaciones = @()

    if ($RAMPct -gt 85) {
        $Estado = "Critico"
        $Recomendaciones += "RAM critica ($RAMPct%). Se recomienda ampliar memoria"
    } elseif ($RAMPct -gt 70) {
        $Estado = "Atencion"
        $Recomendaciones += "RAM elevada ($RAMPct%). Monitorear uso. Considerar ampliacion"
    } else {
        $Recomendaciones += "Equipo operando con recursos suficientes"
    }

    # Alertas adicionales
    if ($ChromeMB -gt 1500) {
        $Recomendaciones += "Chrome consumiendo $ChromeMB MB. Reducir pestanas"
    }
    if ($UptimeDias -gt 15) {
        $Recomendaciones += "Sin reinicio hace $UptimeDias dias. Reiniciar pronto"
    }
    if ($DiscoLibreGB -lt 20) {
        $Recomendaciones += "Disco bajo: $DiscoLibreGB GB libres. Liberar espacio"
    }

    $RecomendacionStr = $Recomendaciones -join " | "

    # =================================================================
    # ENVIAR DATOS A GOOGLE SHEETS
    # =================================================================

    $Body = @{
        Accion         = "diagnostico"
        MACEthernet    = $MACEthernet
        NombreEquipo   = $env:COMPUTERNAME
        RAMTotalGB     = $RAMTotalGB
        RAMUsadaGB     = $RAMUsadaGB
        RAMLibreGB     = $RAMLibreGB
        RAMPct         = $RAMPct
        Top5Procesos   = $Top5Str
        ChromeMB       = $ChromeMB
        ChromeProcs    = $ChromeCount
        DedalusMB      = $DedalusMB
        DedalusProcs   = $DedalusCount
        TotalProcs     = $TotalProcs
        CPUPct         = $CPUPct
        PageFileUsado  = $PageFileUsadoMB
        PageFileTotal  = $PageFileTotalMB
        UptimeDias     = $UptimeDias
        DiscoLibreGB   = "$DiscoLibreGB / $DiscoTotalGB GB"
        Estado         = $Estado
        Recomendacion  = $RecomendacionStr
        FechaReporte   = (Get-Date -Format "dd/MM/yyyy HH:mm")
    } | ConvertTo-Json

    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 30 | Out-Null
            break
        } catch {
            if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 10 -Maximum 30) }
        }
    }

} catch {
    # Silencio total - NUNCA interrumpir Windows
    exit
}
