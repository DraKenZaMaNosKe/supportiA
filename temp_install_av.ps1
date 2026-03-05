# Instalar ESET PROTECT desde la red
$instalador = "\\10.2.1.13\soportefaa\ANTIVIRUS\ESET 2025\PROTECT_Installer_x64_es_CL 2.exe"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  INSTALANDO ESET PROTECT" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Ejecutando: $instalador" -ForegroundColor Yellow
Write-Host "Esto puede tardar varios minutos..." -ForegroundColor Yellow
Write-Host ""

$process = Start-Process -FilePath $instalador -ArgumentList "--silent --accepteula" -Wait -NoNewWindow -PassThru

Start-Sleep -Seconds 5

# Verificar instalacion
$svc = Get-Service -Name "ekrn" -ErrorAction SilentlyContinue
if ($svc) {
    Write-Host ""
    Write-Host "ESET instalado correctamente! Estado: $($svc.Status)" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "Exit code del instalador: $($process.ExitCode)" -ForegroundColor Yellow
    Write-Host "El servicio ekrn aun no aparece. Puede necesitar reinicio." -ForegroundColor Yellow
}
