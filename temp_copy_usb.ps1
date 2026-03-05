$usb = Get-Volume | Where-Object { $_.DriveType -eq 'Removable' -and $_.Size -gt 1GB } | Select-Object -First 1
if ($usb) {
    $letra = $usb.DriveLetter
    Copy-Item "C:\supportiA\ConfiguradorHCG\InstalarReportes.ps1" -Destination "${letra}:\" -Force
    Copy-Item "C:\supportiA\ConfiguradorHCG\CorregirTareasHCG.bat" -Destination "${letra}:\" -Force
    Write-Host "Copiados a ${letra}:\"
    Get-ChildItem "${letra}:\InstalarReportes.ps1", "${letra}:\CorregirTareasHCG.bat" | ForEach-Object { Write-Host "  $($_.Name) - $([math]::Round($_.Length/1KB,1)) KB" }
} else {
    Write-Host "ERROR - No se detecto USB"
}
