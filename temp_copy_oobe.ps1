$usb = Get-Volume | Where-Object { $_.DriveType -eq 'Removable' -and $_.Size -gt 1GB } | Select-Object -First 1
if (-not $usb) {
    Write-Host "ERROR - No se detecto USB" -ForegroundColor Red
    exit
}

$letra = $usb.DriveLetter
Write-Host "USB detectada: ${letra}:\ ($($usb.FileSystemLabel))" -ForegroundColor Green
Write-Host ""

# Copiar autounattend.xml a la raiz (para deteccion automatica)
Copy-Item "C:\supportiA\ConfiguradorHCG\OOBE\autounattend.xml" -Destination "${letra}:\" -Force
Write-Host "  autounattend.xml -> ${letra}:\" -ForegroundColor Gray

# Copiar ConfigOOBE.cmd a la raiz (para acceso rapido desde Shift+F10)
Copy-Item "C:\supportiA\ConfiguradorHCG\OOBE\ConfigOOBE.cmd" -Destination "${letra}:\" -Force
Write-Host "  ConfigOOBE.cmd -> ${letra}:\" -ForegroundColor Gray

# Copiar tambien los otros archivos utiles
Copy-Item "C:\supportiA\ConfiguradorHCG\CorregirTareasHCG.bat" -Destination "${letra}:\" -Force
Write-Host "  CorregirTareasHCG.bat -> ${letra}:\" -ForegroundColor Gray

Copy-Item "C:\supportiA\ConfiguradorHCG\InstalarReportes.ps1" -Destination "${letra}:\" -Force
Write-Host "  InstalarReportes.ps1 -> ${letra}:\" -ForegroundColor Gray

Write-Host ""
Write-Host "Contenido de ${letra}:\:" -ForegroundColor Cyan
Get-ChildItem "${letra}:\" -File | ForEach-Object {
    $tamano = if ($_.Length -gt 1KB) { "$([math]::Round($_.Length/1KB,1)) KB" } else { "$($_.Length) bytes" }
    Write-Host "  $($_.Name)  ($tamano)"
}
