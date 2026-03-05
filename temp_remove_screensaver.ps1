$file = 'C:\supportiA\ConfiguradorHCG\ConfigurarEquipoHCG.ps1'
$lines = Get-Content $file
Write-Host "Total lineas antes: $($lines.Count)"
$before = $lines[0..3436]  # lines 1-3437 (0-indexed)
$after = $lines[3976..($lines.Count-1)]  # line 3977 onwards (Verify-Configuracion)
$result = $before + '' + $after
$result | Set-Content $file -Encoding UTF8
Write-Host "Total lineas despues: $((Get-Content $file).Count)"
Write-Host "Funcion screensaver eliminada correctamente"
