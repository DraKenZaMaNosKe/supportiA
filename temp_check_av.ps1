# Verificar estado de ESET
$svc = Get-Service -Name "ekrn" -ErrorAction SilentlyContinue
if ($svc) {
    Write-Host "ESET instalado - Estado: $($svc.Status)" -ForegroundColor Yellow
} else {
    Write-Host "ESET NO esta instalado" -ForegroundColor Green
}

# Verificar acceso a la red
$ruta = "\\10.2.1.13\soportefaa\ANTIVIRUS\ESET 2025"
if (Test-Path $ruta) {
    Write-Host "Acceso a red: OK" -ForegroundColor Green
    Get-ChildItem $ruta | ForEach-Object { Write-Host "  - $($_.Name)" }
} else {
    Write-Host "Sin acceso a la ruta de red: $ruta" -ForegroundColor Red
}
