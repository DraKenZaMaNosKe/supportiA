# Test de extraccion de datos del hardware

Write-Host ""
Write-Host "  ══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  EXTRACCION DE DATOS DEL HARDWARE - PRUEBA" -ForegroundColor Cyan
Write-Host "  ══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Numero de serie
$Bios = Get-WmiObject Win32_BIOS
Write-Host "  Numero de Serie: " -NoNewline -ForegroundColor Yellow
Write-Host $Bios.SerialNumber -ForegroundColor White

# Fabricante y Modelo
$Sistema = Get-WmiObject Win32_ComputerSystem
Write-Host "  Fabricante:      " -NoNewline -ForegroundColor Yellow
Write-Host $Sistema.Manufacturer -ForegroundColor White
Write-Host "  Modelo:          " -NoNewline -ForegroundColor Yellow
Write-Host $Sistema.Model -ForegroundColor White

# Nombre del equipo
Write-Host "  Nombre Equipo:   " -NoNewline -ForegroundColor Yellow
Write-Host $env:COMPUTERNAME -ForegroundColor White

# UUID
$CS = Get-WmiObject Win32_ComputerSystemProduct
Write-Host "  UUID:            " -NoNewline -ForegroundColor Yellow
Write-Host $CS.UUID -ForegroundColor White

Write-Host ""
Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor Gray

# MAC Addresses
Write-Host ""
Write-Host "  Adaptadores de Red:" -ForegroundColor Cyan

$Adaptadores = Get-WmiObject Win32_NetworkAdapter | Where-Object { $_.MACAddress -and $_.PhysicalAdapter }

foreach ($Adaptador in $Adaptadores) {
    $MAC = $Adaptador.MACAddress -replace ":", ""
    Write-Host "    $($Adaptador.NetConnectionID): " -NoNewline -ForegroundColor Yellow
    Write-Host $MAC -ForegroundColor White
}

Write-Host ""
Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor Gray

# Product Key
Write-Host ""
Write-Host "  Product Key de Windows:" -ForegroundColor Cyan
try {
    $Key = (Get-WmiObject -Query "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService" |
            Where-Object { $_.OA3xOriginalProductKey }).OA3xOriginalProductKey
    if ($Key) {
        Write-Host "    $Key" -ForegroundColor White
    } else {
        Write-Host "    No disponible (licencia por volumen o preinstalada)" -ForegroundColor Gray
    }
} catch {
    Write-Host "    Error al obtener" -ForegroundColor Red
}

Write-Host ""
Write-Host "  ══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""
