# =============================================================================
# REGISTRAR EQUIPOS MANUALMENTE EN GOOGLE SHEETS
# =============================================================================

$GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"

# Forzar TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Datos de los 3 equipos nuevos
$Equipos = @(
    @{
        InvST = "13664"
        Serie = "MZ02WAT1"
        MACEthernet = "2053BD00E458"
        MACWiFi = "14B5CD3004DB"
        ProductKey = "3556359296186"
        FAA = "SI #97"
    },
    @{
        InvST = "13665"
        Serie = "MZ02W5K4"
        MACEthernet = "2053BD00FBD0"
        MACWiFi = "14B5CD31D6AD"
        ProductKey = "3556359102237"
        FAA = "SI #19"
    },
    @{
        InvST = "13666"
        Serie = "MZ02W5PB"
        MACEthernet = "2053BD00FBDB"
        MACWiFi = "14B5CD2FCB57"
        ProductKey = "3556359105105"
        FAA = "SI #92"
    }
)

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  REGISTRAR EQUIPOS EN GOOGLE SHEETS" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

$Registrados = 0
$Errores = 0

foreach ($Equipo in $Equipos) {
    Write-Host "  Registrando equipo: $($Equipo.InvST) - Serie: $($Equipo.Serie)" -ForegroundColor Yellow

    $Body = @{
        Accion = "crear"
        Fecha = (Get-Date -Format "dd/MM/yyyy")
        InvST = $Equipo.InvST
        Serie = $Equipo.Serie
        Marca = "Lenovo"
        Modelo = "ThinkCentre M70s Gen 5"
        Procesador = "Intel Core i5-14500 vPro"
        Nucleos = 14
        RAM = "8"
        Disco = "512"
        DiscoTipo = "SSD"
        Graficos = "Intel UHD 770"
        WiFi = "Wi-Fi 6"
        BT = "5.1"
        SO = "Win 11 Pro"
        MACEthernet = $Equipo.MACEthernet
        MACWiFi = $Equipo.MACWiFi
        ProductKey = $Equipo.ProductKey
        FechaFab = "14/11/2025"
        Garantia = "14/11/2028"
        FAA = $Equipo.FAA  # Enviamos el FAA directamente
    } | ConvertTo-Json

    try {
        $Response = Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 60

        if ($Response.status -eq "OK") {
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |  REGISTRADO EXITOSAMENTE                 |" -ForegroundColor Green
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |  Inventario: $($Equipo.InvST)" -ForegroundColor White
            Write-Host "  |  Serie:      $($Equipo.Serie)" -ForegroundColor White
            Write-Host "  |  Fila:       $($Response.row)" -ForegroundColor Yellow
            Write-Host "  |  FAA:        $($Equipo.FAA)" -ForegroundColor Cyan
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host ""
            $Registrados++
        } else {
            Write-Host "  [WARN] Respuesta inesperada: $($Response | ConvertTo-Json -Compress)" -ForegroundColor Yellow
            $Errores++
        }
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        $Errores++
    }

    Start-Sleep -Milliseconds 1000
}

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  RESUMEN" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  Equipos registrados: $Registrados" -ForegroundColor White
Write-Host "  Errores: $Errores" -ForegroundColor $(if ($Errores -gt 0) { "Red" } else { "White" })
Write-Host ""

Read-Host "  Presiona Enter para salir"
