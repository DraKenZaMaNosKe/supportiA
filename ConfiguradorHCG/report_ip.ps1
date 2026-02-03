# =============================================================================
# HCG - REPORTE AUTOMATICO DE IP
# =============================================================================
# Se ejecuta cada 3 horas y al iniciar sesion de Windows
# Envia la IP actual del equipo a Google Sheets (busca por MAC Ethernet)
# A prueba de errores: si no hay internet, sale silenciosamente
# Con delay aleatorio y reintentos para evitar colisiones
# =============================================================================

$ErrorActionPreference = "SilentlyContinue"

try {
    $GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # --- Delay aleatorio (0-120 seg) para no saturar el servidor ---
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 120)

    # --- Obtener adaptadores de red ---
    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*Ethernet*" } | Select-Object -First 1

    $WiFiAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*Wi-Fi*" -or $_.Name -like "*WiFi*" -or $_.Name -like "*Wireless*" } | Select-Object -First 1

    # --- MAC Ethernet (siempre, incluso si esta desconectado) ---
    $MACEthernet = ""
    if ($EthAdapter) {
        $MACEthernet = ($EthAdapter.MacAddress -replace "-", "").ToUpper()
    }

    # Si no hay MAC Ethernet, no podemos identificar el equipo
    if (-not $MACEthernet) { exit }

    # --- IP Ethernet ---
    $IPEthernet = ""
    if ($EthAdapter -and $EthAdapter.Status -eq "Up") {
        $IPEthernet = (Get-NetIPAddress -InterfaceIndex $EthAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1).IPAddress
    }

    # --- IP WiFi ---
    $IPWiFi = ""
    if ($WiFiAdapter -and $WiFiAdapter.Status -eq "Up") {
        $IPWiFi = (Get-NetIPAddress -InterfaceIndex $WiFiAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1).IPAddress
    }

    # --- SSID WiFi (nombre de la red conectada) ---
    $SSIDWiFi = ""
    if ($WiFiAdapter -and $WiFiAdapter.Status -eq "Up") {
        $NetshOutput = netsh wlan show interfaces 2>$null
        if ($NetshOutput) {
            $SSIDLine = $NetshOutput | Select-String "^\s+SSID\s+:" | Select-Object -First 1
            if ($SSIDLine) {
                $SSIDWiFi = ($SSIDLine.ToString() -replace "^\s+SSID\s+:\s+", "").Trim()
            }
        }
    }

    # Si no hay ninguna IP activa, no tiene caso enviar
    if (-not $IPEthernet -and -not $IPWiFi) { exit }

    # --- Verificar conectividad a internet ---
    $TestOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) {
        $TestOK = Test-Connection -ComputerName "dns.google" -Count 1 -Quiet -ErrorAction SilentlyContinue
    }
    if (-not $TestOK) { exit }

    # --- Enviar datos a Google Sheets (con reintentos) ---
    $Body = @{
        Accion       = "ip"
        MACEthernet  = $MACEthernet
        IPEthernet   = $IPEthernet
        IPWiFi       = $IPWiFi
        SSIDWiFi     = $SSIDWiFi
        NombreEquipo = $env:COMPUTERNAME
        FechaReporte = (Get-Date -Format "dd/MM/yyyy HH:mm")
    } | ConvertTo-Json

    $Enviado = $false
    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 30 | Out-Null
            $Enviado = $true
            break
        } catch {
            if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 10 -Maximum 30) }
        }
    }

} catch {
    # Silencio total - NUNCA interrumpir Windows
    exit
}
