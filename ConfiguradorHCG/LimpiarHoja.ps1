# =============================================================================
# LIMPIAR Y FORMATEAR HOJA DE GOOGLE SHEETS
# =============================================================================

$GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  LIMPIAR Y FORMATEAR GOOGLE SHEETS" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

$Body = @{
    Accion = "limpiar"
} | ConvertTo-Json

Write-Host "  Enviando solicitud de limpieza..." -ForegroundColor Yellow

try {
    $Response = Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 120

    if ($Response.status -eq "OK") {
        Write-Host ""
        Write-Host "  +------------------------------------------+" -ForegroundColor Green
        Write-Host "  |  HOJA LIMPIADA Y FORMATEADA             |" -ForegroundColor Green
        Write-Host "  +------------------------------------------+" -ForegroundColor Green
        Write-Host "  |  Equipos: $($Response.equipos)" -ForegroundColor White
        Write-Host "  |  Filas eliminadas: $($Response.filasEliminadas)" -ForegroundColor White
        Write-Host "  |  $($Response.mensaje)" -ForegroundColor Cyan
        Write-Host "  +------------------------------------------+" -ForegroundColor Green
    } else {
        Write-Host "  [ERROR] $($Response.mensaje)" -ForegroundColor Red
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Read-Host "  Presiona Enter para salir"
