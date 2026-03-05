# Lanzar instalador ESET con GUI
$instalador = "\\10.2.1.13\soportefaa\ANTIVIRUS\ESET 2025\PROTECT_Installer_x64_es_CL 2.exe"

if (Test-Path $instalador) {
    Write-Host "Abriendo instalador ESET..." -ForegroundColor Green
    Start-Process -FilePath $instalador
    Write-Host "El instalador se abrio. Sigue los pasos en pantalla." -ForegroundColor Yellow
} else {
    Write-Host "No se encontro: $instalador" -ForegroundColor Red
}
