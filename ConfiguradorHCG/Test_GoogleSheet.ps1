# =============================================================================
# TEST - Enviar datos de prueba a Google Sheets
# =============================================================================

Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "  TEST DE CONEXION A GOOGLE SHEETS" -ForegroundColor Cyan
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""

# URL del Apps Script (CAMBIAR POR EL NUEVO URL)
$URL = Read-Host "  Pega el URL del Apps Script"

Write-Host ""
Write-Host "  Enviando datos de prueba..." -ForegroundColor Yellow

# Datos de prueba
$DatosPrueba = @{
    Fecha = (Get-Date -Format "dd/MM/yyyy")
    InvST = "99999"
    Serie = "TEST123456"
    MACEthernet = "AABBCCDDEEFF"
    MACWiFi = "112233445566"
    ProductKey = "XXXXX-XXXXX-XXXXX"
    FechaFab = (Get-Date -Format "dd/MM/yyyy")
    Garantia = (Get-Date).AddYears(3).ToString("dd/MM/yyyy")
} | ConvertTo-Json

Write-Host ""
Write-Host "  Datos a enviar:" -ForegroundColor Gray
Write-Host $DatosPrueba -ForegroundColor White
Write-Host ""

try {
    $Response = Invoke-RestMethod -Uri $URL -Method Post -Body $DatosPrueba -ContentType "application/json" -TimeoutSec 30

    Write-Host "  ============================================" -ForegroundColor Green
    Write-Host "  RESPUESTA DEL SERVIDOR:" -ForegroundColor Green
    Write-Host "  ============================================" -ForegroundColor Green
    Write-Host ""
    $Response | ConvertTo-Json | Write-Host
    Write-Host ""

    if ($Response.status -eq "OK") {
        Write-Host "  [OK] Datos enviados correctamente!" -ForegroundColor Green
        Write-Host "  Revisa tu Google Sheet, deberia haber una nueva fila." -ForegroundColor Cyan
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Posibles causas:" -ForegroundColor Yellow
    Write-Host "  - El URL no es correcto" -ForegroundColor Gray
    Write-Host "  - El Apps Script no esta publicado como Web App" -ForegroundColor Gray
    Write-Host "  - Permisos incorrectos en la publicacion" -ForegroundColor Gray
}

Write-Host ""
Read-Host "  Presiona Enter para salir"
